try:
    import pandas as pd
except Exception:
    pd = None
import logging
import mimetypes
import os

from django import forms
from django.conf import settings
from django.contrib import messages
from django.contrib.admin.views.decorators import staff_member_required
from django.contrib.auth import update_session_auth_hash
from django.contrib.auth.decorators import login_required, permission_required
from django.contrib.auth.decorators import user_passes_test
from django.contrib.auth.forms import PasswordChangeForm
from django.http import FileResponse, HttpResponse, JsonResponse
from django.shortcuts import get_object_or_404, redirect, render
from django.template.loader import render_to_string
from django.urls import reverse
from django.views.decorators.http import require_http_methods
from django.core.files.storage import default_storage
from django.utils import timezone
from urllib.parse import quote
import uuid
import os

from cmms.models import Customer_list, Equipment_list

from .forms import ProfileForm
from .models import MasterItem
from .history_utils import log_equipment_history, capture_equipment_snapshot

# Module logger for this views module — use logging instead of print()
logger = logging.getLogger(__name__)


# Simple form for MasterItem
class MasterItemForm(forms.ModelForm):
    class Meta:
        model = MasterItem
        fields = ("category", "code", "label", "description", "active", "order")
        widgets = {
            "category": forms.TextInput(attrs={"class": "form-control"}),
            "code": forms.TextInput(attrs={"class": "form-control"}),
            "label": forms.TextInput(attrs={"class": "form-control"}),
            "description": forms.Textarea(attrs={"class": "form-control", "rows": 3}),
            "active": forms.CheckboxInput(),
            "order": forms.NumberInput(attrs={"class": "form-control"}),
        }

    # allow a hidden parent_id input for prefilling when creating a child via /master/add/?parent_id=NN
    parent_id = forms.IntegerField(
        required=False, widget=forms.HiddenInput(attrs={"id": "mi-parent"})
    )


def favicon(request):
    """Return empty response for favicon requests to avoid 404 noise."""
    return HttpResponse(status=204)


def master_list(request, category=None):
    """List master items; if category provided, filter by it."""
    qs = MasterItem.objects.all()
    categories = (
        MasterItem.objects.values_list("category", flat=True)
        .distinct()
        .order_by("category")
    )

    if category:
        # build tree for the selected category
        nodes = list(
            MasterItem.objects.filter(category=category).order_by("order", "label")
        )
        node_map = {n.id: {"item": n, "children": []} for n in nodes}
        roots = []
        for n in nodes:
            if n.parent_id and n.parent_id in node_map:
                node_map[n.parent_id]["children"].append(node_map[n.id])
            else:
                roots.append(node_map[n.id])

        return render(
            request,
            "master/master_list.html",
            {
                "items": None,
                "categories": categories,
                "active_category": category,
                "tree_roots": roots,
            },
        )

    return render(
        request,
        "master/master_list.html",
        {"items": qs, "categories": categories, "active_category": category},
    )


@staff_member_required
@require_http_methods(["GET", "POST"])
def master_import_excel(request):
    """Upload an Excel (XLSX) file to bulk import MasterItem rows.

    Expected sheet layout: header row with columns: category, code, label, description, active, order
    """
    import openpyxl

    from .models import MasterItem

    result = {"created": 0, "updated": 0, "skipped": 0, "errors": []}

    if request.method == "POST":
        f = request.FILES.get("file")
        if not f:
            messages.error(request, "ไม่พบไฟล์ที่อัปโหลด")
            return redirect("master_import_excel")

        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            ws = wb.active
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                messages.error(request, "ไฟล์ว่างเปล่า")
                return redirect("master_import_excel")
            headers = [str(h).strip() if h is not None else "" for h in rows[0]]
            # normalize header names
            headers = [h.lower() for h in headers]
            expected = [
                "category",
                "code",
                "label",
                "description",
                "active",
                "order",
                "parent_category",
                "parent_code",
            ]
            # collect any meta.* headers (e.g. meta.floor, meta.capacity)
            meta_headers = [h for h in headers if h.startswith("meta.")]
            # find indices
            idx = {h: (headers.index(h) if h in headers else None) for h in expected}
            # indices for meta headers
            meta_idx = {h: headers.index(h) for h in meta_headers}

            for rnum, row in enumerate(rows[1:], start=2):
                try:
                    vals = {
                        h: (
                            row[idx[h]]
                            if idx[h] is not None and idx[h] < len(row)
                            else None
                        )
                        for h in expected
                    }
                    # collect meta raw values per-row
                    meta_raw = {}
                    for mh, mpos in meta_idx.items():
                        if mpos is not None and mpos < len(row):
                            # mh like 'meta.floor' -> key 'floor'
                            key = mh.split(".", 1)[1]
                            meta_raw[key] = row[mpos]
                    category = (
                        str(vals["category"]).strip()
                        if vals["category"] is not None
                        else ""
                    )
                    label = (
                        str(vals["label"]).strip() if vals["label"] is not None else ""
                    )
                    code = str(vals["code"]).strip() if vals["code"] is not None else ""
                    description = (
                        str(vals["description"]).strip()
                        if vals["description"] is not None
                        else ""
                    )
                    active = (
                        str(vals["active"]).strip().lower()
                        in ("1", "true", "yes", "y", "t")
                        if vals["active"] is not None
                        else True
                    )
                    order = 0
                    try:
                        order = (
                            int(vals["order"])
                            if vals["order"] is not None and vals["order"] != ""
                            else 0
                        )
                    except Exception:
                        order = 0

                    if not category or not label:
                        result["skipped"] += 1
                        continue

                    # If there are meta columns, attempt to coerce/validate them using MasterCategory definitions
                    meta_values = {}
                    try:
                        from .models import MasterCategory

                        cat = MasterCategory.objects.filter(key=category).first()
                        if cat and meta_raw:
                            # build a map of field defs by name
                            fdefs = {f.name: f for f in cat.fields.all()}
                            for k, raw in meta_raw.items():
                                raw_value = "" if raw is None else str(raw).strip()
                                if not raw_value:
                                    meta_values[k] = None
                                    continue
                                fdef = fdefs.get(k)
                                if fdef:
                                    try:
                                        if fdef.field_type == "integer":
                                            meta_values[k] = int(raw_value)
                                        elif fdef.field_type == "decimal":
                                            from decimal import Decimal

                                            meta_values[k] = Decimal(raw_value)
                                        elif fdef.field_type == "boolean":
                                            meta_values[k] = raw_value.lower() in (
                                                "1",
                                                "true",
                                                "on",
                                                "yes",
                                            )
                                        elif fdef.field_type == "date":
                                            import datetime

                                            meta_values[k] = (
                                                datetime.date.fromisoformat(
                                                    raw_value
                                                ).isoformat()
                                            )
                                        else:
                                            # string, text, select
                                            meta_values[k] = raw_value
                                    except Exception as e:
                                        result["errors"].append(
                                            f"Row {rnum} meta.{k}: {e}"
                                        )
                                        # store raw fallback
                                        meta_values[k] = raw_value
                                else:
                                    # no fdef found — store raw value but warn
                                    meta_values[k] = raw_value
                                    result["errors"].append(
                                        f'Row {rnum}: meta field "{k}" not defined for category "{category}"'
                                    )
                    except Exception:
                        # Fail-safe: ignore meta processing errors and continue
                        meta_values = {}

                    # lookup existing by (category, code) if code present, else by (category,label)
                    existing = None
                    if code:
                        existing = MasterItem.objects.filter(
                            category=category, code=code
                        ).first()
                    if not existing:
                        existing = MasterItem.objects.filter(
                            category=category, label=label
                        ).first()

                    if existing:
                        # update
                        existing.label = label
                        existing.description = description
                        existing.active = active
                        existing.order = order
                        if code:
                            existing.code = code
                        # attach meta if available
                        try:
                            if meta_values:
                                existing.meta = meta_values
                        except Exception:
                            pass
                        existing.save()
                        result["updated"] += 1
                    else:
                        # create without parent first
                        mi = MasterItem.objects.create(
                            category=category,
                            code=code,
                            label=label,
                            description=description,
                            active=active,
                            order=order,
                        )
                        try:
                            if meta_values:
                                mi.meta = meta_values
                                mi.save()
                        except Exception:
                            pass
                        result["created"] += 1

                except Exception as e:
                    result["errors"].append(f"Row {rnum}: {e}")

            # second pass: link parents if parent_code provided
            linked = 0
            for rnum, row in enumerate(rows[1:], start=2):
                try:
                    # extract parent info
                    parent_category = None
                    parent_code = None
                    if idx.get("parent_category") is not None and idx.get(
                        "parent_category"
                    ) < len(row):
                        parent_category = row[idx["parent_category"]]
                    if idx.get("parent_code") is not None and idx.get(
                        "parent_code"
                    ) < len(row):
                        parent_code = row[idx["parent_code"]]
                    if not parent_code:
                        continue
                    # find child
                    code = (
                        row[idx["code"]]
                        if idx.get("code") is not None and idx["code"] < len(row)
                        else None
                    )
                    label = (
                        row[idx["label"]]
                        if idx.get("label") is not None and idx["label"] < len(row)
                        else None
                    )
                    category = (
                        row[idx["category"]]
                        if idx.get("category") is not None
                        and idx["category"] < len(row)
                        else None
                    )
                    child = None
                    if code:
                        child = MasterItem.objects.filter(
                            category=category, code=str(code).strip()
                        ).first()
                    if not child and label:
                        child = MasterItem.objects.filter(
                            category=category, label=str(label).strip()
                        ).first()
                    if not child:
                        continue
                    parent = None
                    if parent_category:
                        parent = MasterItem.objects.filter(
                            category=str(parent_category).strip(),
                            code=str(parent_code).strip(),
                        ).first()
                    if not parent:
                        parent = MasterItem.objects.filter(
                            code=str(parent_code).strip()
                        ).first()
                    if parent:
                        child.parent = parent
                        child.save()
                        linked += 1
                except Exception as e:
                    result["errors"].append(f"Link Row {rnum}: {e}")

            messages.success(
                request,
                f"Import finished — created={result['created']} updated={result['updated']} skipped={result['skipped']} linked={linked} errors={len(result['errors'])}",
            )
            if result["errors"]:
                for e in result["errors"][:10]:
                    messages.error(request, e)
            return redirect("master_list")

        except Exception as e:
            messages.error(request, f"Error reading file: {e}")
            return redirect("master_import_excel")

    # GET
    # Provide list of master categories for template download dropdown
    try:
        from .models import MasterCategory

        categories = list(
            MasterCategory.objects.all().order_by("key").values("key", "label")
        )
    except Exception:
        categories = []

    return render(request, "master/master_import.html", {"categories": categories})


@staff_member_required
def master_download_template(request):
    """Generate and return an Excel .xlsx template file for MasterItem import."""
    from django.http import HttpResponse
    from openpyxl.workbook import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "master_template"

    # Base headers (import supports additional columns named meta.<field_name>)
    headers = [
        "category",
        "code",
        "label",
        "description",
        "active",
        "order",
        "parent_category",
        "parent_code",
    ]

    # If user requested a category-specific template, include meta.<field> headers for that category
    requested_category = (request.GET.get("category") or "").strip()
    cat_obj = None
    try:
        from .models import MasterCategory

        if requested_category:
            cat_obj = MasterCategory.objects.filter(key=requested_category).first()
            if cat_obj:
                # include meta headers for this category
                for f in list(cat_obj.fields.order_by("order")):
                    headers.append(f"meta.{f.name}")
                # set a helpful sheet title (Excel limits 31 chars)
                ws.title = f"master_{requested_category}"[:31]
    except Exception:
        cat_obj = None

    ws.append(headers)

    # sample rows — pre-fill category column if category-specific template requested
    sample_category = requested_category if cat_obj else "customers"

    # Build example meta values for the selected category (if any)
    meta_examples = []
    if cat_obj:
        for f in list(cat_obj.fields.order_by("order")):
            ex = ""
            try:
                # prefer default_value if provided
                if getattr(f, "default_value", None):
                    ex = str(f.default_value)
                elif f.field_type == "integer":
                    ex = "1"
                elif f.field_type == "decimal":
                    ex = "1.5"
                elif f.field_type == "boolean":
                    ex = "true"
                elif f.field_type == "date":
                    import datetime

                    ex = datetime.date.today().isoformat()
                elif f.field_type == "select":
                    # try to use first choice value
                    if f.choices:
                        try:
                            choices = f.choices
                            if isinstance(choices, list) and len(choices) > 0:
                                first = choices[0]
                                if isinstance(first, dict):
                                    ex = str(first.get("value", ""))
                                else:
                                    ex = str(first)
                        except Exception:
                            ex = ""
                    else:
                        ex = ""
                else:
                    ex = "example"
            except Exception:
                ex = ""
            meta_examples.append(ex)

    # If we have existing MasterItem rows for this category, use them as examples
    used_examples = []
    if cat_obj:
        try:
            from .models import MasterItem

            items = list(
                MasterItem.objects.filter(category=cat_obj.key).order_by("order")[:3]
            )
            for ex_item in items:
                row = [
                    cat_obj.key,
                    ex_item.code or "",
                    ex_item.label or "",
                    ex_item.description or "",
                    "true" if ex_item.active else "false",
                    ex_item.order or 0,
                    "",
                    "",
                ]
                # append meta values in the same order as field definitions
                for f in list(cat_obj.fields.order_by("order")):
                    val = ""
                    try:
                        if getattr(ex_item, "meta", None) and isinstance(
                            ex_item.meta, dict
                        ):
                            mv = ex_item.meta.get(f.name)
                            if mv is not None:
                                # format dates as ISO if necessary
                                if f.field_type == "date":
                                    try:
                                        # assume stored as ISO string
                                        val = str(mv)
                                    except Exception:
                                        val = str(mv)
                                else:
                                    val = str(mv)
                    except Exception:
                        val = ""
                    row.append(val)
                used_examples.append(row)
        except Exception:
            used_examples = []

    # If no real items, fall back to static example rows + meta_examples
    if used_examples:
        rows_to_write = used_examples
    else:
        base_row1 = [
            sample_category,
            "cntu",
            "คณะแพทยศาสตร์ มหาวิทยาลัยเชียงใหม่",
            "Imported customer",
            "true",
            10,
            "",
            "",
        ]
        base_row2 = [
            sample_category if cat_obj else "division",
            "cntu_div1",
            "Division 1",
            "",
            "true",
            20,
            "customers",
            "cntu",
        ]
        base_row3 = [
            sample_category if cat_obj else "department",
            "cntu_dep1",
            "ภาควิชา X",
            "",
            "true",
            30,
            "division",
            "cntu_div1",
        ]
        if meta_examples:
            base_row1.extend(meta_examples)
            base_row2.extend(meta_examples)
            base_row3.extend(meta_examples)
        rows_to_write = [base_row1, base_row2, base_row3]

    # ensure sample rows match header length and append
    for r in rows_to_write:
        while len(r) < len(headers):
            r.append("")
        ws.append(r)

    # Add a second sheet that documents meta fields for each MasterCategory (if any)
    try:
        from .models import MasterCategory

        meta_ws = wb.create_sheet(title="meta_fields")
        meta_ws.append(
            [
                "category_key",
                "field_name",
                "label",
                "field_type",
                "required",
                "choices",
                "help_text",
            ]
        )
        for cat in MasterCategory.objects.all().order_by("key"):
            # if no fields, still add a row describing the category
            fields = list(cat.fields.order_by("order"))
            if not fields:
                meta_ws.append([cat.key, "", "", "", "", "", ""])
                continue
            for f in fields:
                # choices may be a JSON/list — convert to readable string
                choices_str = ""
                try:
                    if f.choices:
                        import json

                        choices_str = json.dumps(f.choices, ensure_ascii=False)
                except Exception:
                    choices_str = str(f.choices)
                meta_ws.append(
                    [
                        cat.key,
                        f.name,
                        f.label,
                        f.field_type,
                        "yes" if f.required else "no",
                        choices_str,
                        f.help_text or "",
                    ]
                )
    except Exception:
        # ignore metadata sheet generation failures — still return base template
        pass

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    filename = (
        f"master_template_{requested_category}.xlsx"
        if requested_category
        else "master_template.xlsx"
    )
    response["Content-Disposition"] = f"attachment; filename={filename}"
    wb.save(response)
    return response


@require_http_methods(["GET", "POST"])
def master_create(request):
    from .models import MasterCategory

    if request.method == "POST":
        form = MasterItemForm(request.POST)

        # Load dynamic fields for validation
        category_key = request.POST.get("category", "").strip()
        field_defs = []
        meta_errors = {}
        meta_values = {}

        try:
            cat = MasterCategory.objects.get(key=category_key)
            field_defs = list(cat.fields.order_by("order"))
        except MasterCategory.DoesNotExist:
            field_defs = []

        # Validate and coerce meta fields
        for fdef in field_defs:
            field_key = f"meta.{fdef.name}"
            raw_value = request.POST.get(field_key, "").strip()

            # Required check
            if fdef.required and not raw_value:
                meta_errors[fdef.name] = f"{fdef.label} เป็นฟิลด์ที่จำเป็น"
                continue

            # Type coercion & validation
            try:
                if fdef.field_type == "integer":
                    meta_values[fdef.name] = int(raw_value) if raw_value else None
                elif fdef.field_type == "decimal":
                    from decimal import Decimal

                    meta_values[fdef.name] = Decimal(raw_value) if raw_value else None
                elif fdef.field_type == "boolean":
                    meta_values[fdef.name] = raw_value.lower() in (
                        "1",
                        "true",
                        "on",
                        "yes",
                    )
                elif fdef.field_type == "date":
                    if raw_value:
                        import datetime

                        meta_values[fdef.name] = datetime.date.fromisoformat(
                            raw_value
                        ).isoformat()
                    else:
                        meta_values[fdef.name] = None
                elif fdef.field_type == "select":
                    # Validate against choices
                    if raw_value and fdef.choices:
                        valid_values = []
                        if isinstance(fdef.choices, list):
                            for choice in fdef.choices:
                                if isinstance(choice, dict):
                                    valid_values.append(str(choice.get("value", "")))
                                else:
                                    valid_values.append(str(choice))
                        if raw_value not in valid_values:
                            meta_errors[fdef.name] = (
                                f"ค่า '{raw_value}' ไม่ถูกต้องสำหรับ {fdef.label}"
                            )
                        else:
                            meta_values[fdef.name] = raw_value
                    else:
                        meta_values[fdef.name] = raw_value if raw_value else None
                else:
                    # string, text
                    meta_values[fdef.name] = raw_value if raw_value else None
            except (ValueError, TypeError) as e:
                meta_errors[fdef.name] = f"ค่าไม่ถูกต้อง: {e}"

        if form.is_valid() and not meta_errors:
            # Save base fields
            parent_id = form.cleaned_data.get("parent_id")
            item = form.save(commit=False)
            if parent_id:
                try:
                    item.parent_id = int(parent_id)
                except Exception:
                    item.parent = None
            # Save meta
            item.meta = meta_values if meta_values else None
            item.save()
            form.save_m2m()
            messages.success(request, "เพิ่มข้อมูล Master สำเร็จ")
            return redirect("master_list")
        else:
            # Re-render with errors
            context = {
                "form": form,
                "create": True,
                "field_defs": field_defs,
                "meta_errors": meta_errors,
                "meta_values": meta_values,
            }
            return render(request, "master/master_form.html", context)
    else:
        # GET: allow pre-filling parent via ?parent_id= and category via ?category=
        initial = {}
        parent_id = request.GET.get("parent_id")
        parent_label = None
        if parent_id:
            try:
                parent_obj = MasterItem.objects.get(pk=int(parent_id))
                initial["parent_id"] = parent_obj.pk
                parent_label = f"{parent_obj.label}{(' ('+parent_obj.code+')') if parent_obj.code else ''}"
            except Exception:
                parent_label = None

        # Pre-fill category if provided
        category_key = request.GET.get("category", "").strip()
        if category_key:
            initial["category"] = category_key

        form = MasterItemForm(initial=initial)

        # Load field definitions for the category
        field_defs = []
        if category_key:
            try:
                cat = MasterCategory.objects.get(key=category_key)
                field_defs = list(cat.fields.order_by("order"))
            except MasterCategory.DoesNotExist:
                field_defs = []

        context = {
            "form": form,
            "create": True,
            "parent_label": parent_label,
            "parent_id": initial.get("parent_id", ""),
            "field_defs": field_defs,
            "meta_errors": {},
            "meta_values": {},
        }
        return render(request, "master/master_form.html", context)


@require_http_methods(["GET", "POST"])
def master_edit(request, pk):
    from .models import MasterCategory

    item = get_object_or_404(MasterItem, pk=pk)

    if request.method == "POST":
        form = MasterItemForm(request.POST, instance=item)

        # Load dynamic fields for validation
        category_key = request.POST.get("category", "").strip()
        field_defs = []
        meta_errors = {}
        meta_values = {}

        try:
            cat = MasterCategory.objects.get(key=category_key)
            field_defs = list(cat.fields.order_by("order"))
        except MasterCategory.DoesNotExist:
            field_defs = []

        # Validate and coerce meta fields
        for fdef in field_defs:
            field_key = f"meta.{fdef.name}"
            raw_value = request.POST.get(field_key, "").strip()

            # Required check
            if fdef.required and not raw_value:
                meta_errors[fdef.name] = f"{fdef.label} เป็นฟิลด์ที่จำเป็น"
                continue

            # Type coercion & validation
            try:
                if fdef.field_type == "integer":
                    meta_values[fdef.name] = int(raw_value) if raw_value else None
                elif fdef.field_type == "decimal":
                    from decimal import Decimal

                    meta_values[fdef.name] = Decimal(raw_value) if raw_value else None
                elif fdef.field_type == "boolean":
                    meta_values[fdef.name] = raw_value.lower() in (
                        "1",
                        "true",
                        "on",
                        "yes",
                    )
                elif fdef.field_type == "date":
                    if raw_value:
                        import datetime

                        meta_values[fdef.name] = datetime.date.fromisoformat(
                            raw_value
                        ).isoformat()
                    else:
                        meta_values[fdef.name] = None
                elif fdef.field_type == "select":
                    # Validate against choices
                    if raw_value and fdef.choices:
                        valid_values = []
                        if isinstance(fdef.choices, list):
                            for choice in fdef.choices:
                                if isinstance(choice, dict):
                                    valid_values.append(str(choice.get("value", "")))
                                else:
                                    valid_values.append(str(choice))
                        if raw_value not in valid_values:
                            meta_errors[fdef.name] = (
                                f"ค่า '{raw_value}' ไม่ถูกต้องสำหรับ {fdef.label}"
                            )
                        else:
                            meta_values[fdef.name] = raw_value
                    else:
                        meta_values[fdef.name] = raw_value if raw_value else None
                else:
                    # string, text
                    meta_values[fdef.name] = raw_value if raw_value else None
            except (ValueError, TypeError) as e:
                meta_errors[fdef.name] = f"ค่าไม่ถูกต้อง: {e}"

        if form.is_valid() and not meta_errors:
            # Save with meta
            item = form.save(commit=False)
            item.meta = meta_values if meta_values else None
            item.save()
            form.save_m2m()
            messages.success(request, "แก้ไขข้อมูล Master สำเร็จ")
            return redirect("master_list")
        else:
            context = {
                "form": form,
                "create": False,
                "item": item,
                "field_defs": field_defs,
                "meta_errors": meta_errors,
                "meta_values": meta_values,
            }
            return render(request, "master/master_form.html", context)
    else:
        form = MasterItemForm(instance=item)

        # Load field definitions for the item's category
        field_defs = []
        try:
            cat = MasterCategory.objects.get(key=item.category)
            field_defs = list(cat.fields.order_by("order"))
        except MasterCategory.DoesNotExist:
            field_defs = []

        context = {
            "form": form,
            "create": False,
            "item": item,
            "field_defs": field_defs,
            "meta_errors": {},
            "meta_values": item.meta or {},
        }
        return render(request, "master/master_form.html", context)


@require_http_methods(["POST"])
def master_delete(request, pk):
    item = get_object_or_404(MasterItem, pk=pk)
    category = item.category  # Save category before deleting
    item.delete()
    
    # Try to redirect back to category page if category exists
    if category:
        try:
            return redirect("master_list_by_category", category=category)
        except:
            pass
    
    # Fallback: use referer or master_list
    referer = request.META.get('HTTP_REFERER')
    if referer and '/master/' in referer:
        # Clean up referer URL to avoid issues
        from urllib.parse import urlparse
        parsed = urlparse(referer)
        if parsed.path and parsed.path != request.path:
            return redirect(parsed.path)
    
    return redirect("master_list")


# Tools for inspecting templates (staff-only)
@staff_member_required
def templates_index(request):
    """List template files under the app's templates directory."""
    templates_dir = os.path.join(settings.BASE_DIR, "cmms", "templates")
    files = []
    for root, dirs, filenames in os.walk(templates_dir):
        for fn in filenames:
            if fn.endswith(".html"):
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, templates_dir).replace("\\", "/")
                files.append(rel)
    files.sort()

    # group by top-level folder (or 'top-level' for templates at root)
    from collections import OrderedDict, defaultdict

    groups = defaultdict(list)

    # small mapping of group display names and descriptions (Thai)
    group_display = {
        "tools": "เครื่องมือ (Tools)",
        "equipment": "อุปกรณ์ (Equipment)",
        "maintenance": "การบำรุงรักษา (Maintenance)",
        "master": "รายการ Master (Master data)",
        "accounts": "บัญชีและการเข้าถึง (Accounts)",
    }

    group_descriptions = {
        "tools": "หน้าสำหรับนักพัฒนาและผู้ดูแลระบบ ใช้ดู/เรนเดอร์/ดูซอร์สของเทมเพลต",
        "equipment": "เทมเพลตที่เกี่ยวข้องกับข้อมูลและหน้าอุปกรณ์",
        "maintenance": "หน้าที่เกี่ยวกับแผนการบำรุงรักษา นัดหมาย และการจัดการช่าง",
        "master": "เทมเพลตสำหรับจัดการ master lists / ตัวเลือก",
    }

    # per-template short descriptions (Thai). Add common ones; others will get a generated friendly label
    per_template_desc = {
        "tools/templates_index.html": "หน้ารายการเทมเพลต พร้อมลิงก์ดูซอร์สและตัวอย่างการเรนเดอร์",
        "equipment/maintenance_schedule.html": "หน้าแผนการบำรุงรักษา (เลือกวันที่นัดหมายสำหรับอุปกรณ์ที่ใกล้ครบกำหนด)",
        "maintenance/technicians_list.html": "รายการช่างและข้อมูลพื้นฐาน",
        "maintenance/technician_availability_manage.html": "จัดการตารางการมาทำงานของช่างในแต่ละวัน",
        "master/master_list.html": "แสดงรายการ master items เป็นลิสต์หรือต้นไม้",
        "master/tree_manage.html": "หน้าจัดการ master แบบโครงสร้างต้นไม้ (tree)",
        "master/master_form.html": "ฟอร์มเพิ่ม/แก้ไข master item",
    }

    for rel in files:
        if "/" in rel:
            grp = rel.split("/")[0]
        else:
            grp = "top-level"
        # friendly label
        label = os.path.basename(rel)
        desc = per_template_desc.get(rel, "")
        if not desc:
            # generate a readable description from filename
            name = os.path.splitext(label)[0].replace("_", " ").replace("-", " ")
            desc = f"ไฟล์เทมเพลต: {name}"
        groups[grp].append({"path": rel, "label": label, "description": desc})

    # order groups: top-level first, then alphabetical
    ordered = OrderedDict()

    def _make_group_entry(gk, items):
        name = group_display.get(gk, gk)
        desc = group_descriptions.get(gk, "")
        return {"name": name, "description": desc, "items": items}

    if "top-level" in groups:
        ordered["top-level"] = _make_group_entry("top-level", groups.pop("top-level"))
    for k in sorted(groups.keys()):
        ordered[k] = _make_group_entry(k, groups[k])

    return render(
        request,
        "tools/templates_index.html",
        {
            "template_groups": ordered,
        },
    )


@staff_member_required
def templates_index_csv(request):
    """Return a CSV file listing templates with group and description."""
    import csv
    from io import StringIO

    templates_dir = os.path.join(settings.BASE_DIR, "cmms", "templates")
    files = []
    for root, dirs, filenames in os.walk(templates_dir):
        for fn in filenames:
            if fn.endswith(".html"):
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, templates_dir).replace("\\", "/")
                files.append(rel)
    files.sort()

    # group mapping same as templates_index
    group_display = {
        "tools": "เครื่องมือ (Tools)",
        "equipment": "อุปกรณ์ (Equipment)",
        "maintenance": "การบำรุงรักษา (Maintenance)",
        "master": "รายการ Master (Master data)",
        "accounts": "บัญชีและการเข้าถึง (Accounts)",
    }
    group_descriptions = {
        "tools": "หน้าสำหรับนักพัฒนาและผู้ดูแลระบบ ใช้ดู/เรนเดอร์/ดูซอร์สของเทมเพลต",
        "equipment": "เทมเพลตที่เกี่ยวข้องกับข้อมูลและหน้าอุปกรณ์",
        "maintenance": "หน้าที่เกี่ยวกับแผนการบำรุงรักษา นัดหมาย และการจัดการช่าง",
        "master": "เทมเพลตสำหรับจัดการ master lists / ตัวเลือก",
    }

    out = StringIO()
    writer = csv.writer(out)
    writer.writerow(["path", "group", "group_display", "group_description"])
    for rel in files:
        grp = rel.split("/")[0] if "/" in rel else "top-level"
        writer.writerow(
            [rel, grp, group_display.get(grp, grp), group_descriptions.get(grp, "")]
        )

    response = HttpResponse(out.getvalue(), content_type="text/csv")
    response["Content-Disposition"] = "attachment; filename=templates_list.csv"
    return response


@staff_member_required
def templates_index_group_csv(request, group_key):
    """Return CSV of templates for a specific top-level group."""
    import csv
    from io import StringIO

    templates_dir = os.path.join(settings.BASE_DIR, "cmms", "templates")
    files = []
    for root, dirs, filenames in os.walk(templates_dir):
        for fn in filenames:
            if fn.endswith(".html"):
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, templates_dir).replace("\\", "/")
                files.append(rel)
    files.sort()

    out = StringIO()
    writer = csv.writer(out)
    writer.writerow(["path"])
    for rel in files:
        grp = rel.split("/")[0] if "/" in rel else "top-level"
        if grp == group_key:
            writer.writerow([rel])

    response = HttpResponse(out.getvalue(), content_type="text/csv")
    response["Content-Disposition"] = f"attachment; filename=templates_{group_key}.csv"
    return response


@staff_member_required
def template_source(request):
    path = request.GET.get("path", "")
    templates_dir = os.path.join(settings.BASE_DIR, "cmms", "templates")
    target = os.path.normpath(os.path.join(templates_dir, path))
    # Prevent path traversal
    if not target.startswith(os.path.normpath(templates_dir)) or not os.path.exists(
        target
    ):
        return HttpResponse("File not found or invalid path", status=404)
    with open(target, "r", encoding="utf-8") as f:
        src = f.read()
    return render(request, "tools/template_source.html", {"path": path, "source": src})


@staff_member_required
def template_render(request):
    # render template with minimal safe context
    path = request.GET.get("path", "")
    # path is relative to cmms/templates
    template_name = path.replace("\\", "/")
    try:
        # render with empty context (but include messages, user)
        ctx = {"user": request.user}
        html = render_to_string(template_name, ctx, request=request)
        return render(
            request, "tools/template_render.html", {"path": path, "rendered": html}
        )
    except Exception as e:
        return render(
            request, "tools/template_render.html", {"path": path, "error": str(e)}
        )


# Create your views here.
def index(request):
    # CMMS dashboard statistics
    from datetime import timedelta

    from django.utils import timezone

    try:
        equipment_count = Equipment_list.objects.count()
    except Exception:
        equipment_count = 0

    try:
        from cmms.models import WorkOrder

        workorder_qs = WorkOrder.objects.all()
        workorder_count = workorder_qs.count()
        # counts by status
        status_counts = {
            "open": workorder_qs.filter(status="open").count(),
            "assigned": workorder_qs.filter(status="assigned").count(),
            "in_progress": workorder_qs.filter(status="in_progress").count(),
            "completed": workorder_qs.filter(status="completed").count(),
            "closed": workorder_qs.filter(status="closed").count(),
        }
        # recently created workorders (last 7 days)
        recent_workorders = workorder_qs.order_by("-reported_at")[:6]
        # pending (open/assigned/in_progress)
        pending_count = workorder_qs.filter(
            status__in=["open", "assigned", "in_progress"]
        ).count()
        # today's workorders
        today = timezone.localtime().date()
        today_count = workorder_qs.filter(reported_at__date=today).count()
    except Exception:
        workorder_count = 0
        status_counts = {
            "open": 0,
            "assigned": 0,
            "in_progress": 0,
            "completed": 0,
            "closed": 0,
        }
        recent_workorders = []
        pending_count = 0
        today_count = 0

    # upcoming PM due within next 30 days
    try:
        now = timezone.localtime()
        soon = now + timedelta(days=30)
        due_pm_qs = Equipment_list.objects.filter(
            equipment_pm_due__gte=now.date(), equipment_pm_due__lte=soon.date()
        ).order_by("equipment_pm_due")[:6]
        due_pm_count = due_pm_qs.count()
    except Exception:
        due_pm_qs = []
        due_pm_count = 0

    context = {
        "equipment_count": equipment_count,
        "workorder_count": workorder_count,
        "status_counts": status_counts,
        "recent_workorders": recent_workorders,
        "due_pm_count": due_pm_count,
        "due_pm_list": due_pm_qs,
        "pending_count": pending_count if "pending_count" in locals() else 0,
        "today_count": today_count if "today_count" in locals() else 0,
    }
    return render(request, "index.html", context)


def api_dashboard_stats(request):
    """Return JSON with basic dashboard stats for AJAX polling."""
    try:
        equipment_count = Equipment_list.objects.count()
    except Exception:
        equipment_count = 0

    try:
        from cmms.models import WorkOrder

        wqs = WorkOrder.objects.all()
        workorder_count = wqs.count()
        status_counts = {
            "open": wqs.filter(status="open").count(),
            "assigned": wqs.filter(status="assigned").count(),
            "in_progress": wqs.filter(status="in_progress").count(),
            "completed": wqs.filter(status="completed").count(),
            "closed": wqs.filter(status="closed").count(),
        }
        pending_count = wqs.filter(
            status__in=["open", "assigned", "in_progress"]
        ).count()
        from django.utils import timezone

        today = timezone.localtime().date()
        today_count = wqs.filter(reported_at__date=today).count()
    except Exception:
        workorder_count = 0
        status_counts = {
            "open": 0,
            "assigned": 0,
            "in_progress": 0,
            "completed": 0,
            "closed": 0,
        }
        pending_count = 0
        today_count = 0

    return JsonResponse(
        {
            "equipment_count": equipment_count,
            "workorder_count": workorder_count,
            "status_counts": status_counts,
            "pending_count": pending_count,
            "today_count": today_count,
        }
    )


@login_required
def api_search_equipment(request):
    """AJAX endpoint for Select2: ?q=term -> { results: [{id, text}, ...] }

    Returns equipment id and name as text.
    """
    q = (request.GET.get("q") or "").strip()
    # optional filters from UI
    requires_pm = request.GET.get("requires_pm")
    requires_cal = request.GET.get("requires_cal")
    qs = Equipment_list.objects.all()
    if q:
        from django.db.models import Q

        qs = qs.filter(
            Q(equipment_id__icontains=q)
            | Q(equipment_name_TH__icontains=q)
            | Q(equipment_name_EN__icontains=q)
        )
    items = []
    # apply filters after initial q-filter
    if requires_pm in ("1", "true", "True"):
        qs = qs.filter(requires_pm=True)
    if requires_cal in ("1", "true", "True"):
        qs = qs.filter(requires_cal=True)

    for e in qs.order_by("equipment_id")[:25]:
        # The detailed owner/user fields (unit/section/department/division)
        # were removed from the model in favor of a single
        # `equipment_user_customer`/`equipment_owner_customer` master label.
        # For backward compatibility, surface the customer label into the
        # user_* keys so clients expecting those keys still receive a value.
        user_customer = getattr(e, "equipment_user_customer", "") or ""
        items.append(
            {
                "id": e.id,
                "text": f"{e.equipment_id} - {e.equipment_name_TH}",
                "equipment_id": e.equipment_id,
                "equipment_name_TH": e.equipment_name_TH,
                "equipment_name_EN": e.equipment_name_EN,
                "brand": e.equipment_brand,
                "model": e.equipment_model,
                "user_unit": user_customer,
                "user_section": user_customer,
                "user_department": user_customer,
                "user_division": user_customer,
                "user_customer": user_customer,
                "location": (
                    getattr(e, "location", "") if hasattr(e, "location") else ""
                ),
                "note": (
                    getattr(e, "equipment_note", "")
                    if hasattr(e, "equipment_note")
                    else ""
                ),
                "requires_pm": bool(e.requires_pm),
                "requires_cal": bool(e.requires_cal),
                "pm_due": (
                    e.equipment_pm_due.isoformat()
                    if getattr(e, "equipment_pm_due", None)
                    else None
                ),
                "cal_due": (
                    e.equipment_cal_due.isoformat()
                    if getattr(e, "equipment_cal_due", None)
                    else None
                ),
            }
        )
    return JsonResponse({"results": items})


@login_required
def api_search_users(request):
    """AJAX endpoint for Select2 users search."""
    q = (request.GET.get("q") or "").strip()
    from django.contrib.auth import get_user_model

    User = get_user_model()
    qs = User.objects.all()
    if q:
        from django.db.models import Q

        qs = qs.filter(
            Q(username__icontains=q)
            | Q(first_name__icontains=q)
            | Q(last_name__icontains=q)
            | Q(email__icontains=q)
        )
    items = []
    for u in qs.order_by("username")[:25]:
        label = u.get_full_name() or u.username
        items.append({"id": u.id, "text": label})
    return JsonResponse({"results": items})


def api_master_items(request):
    """Return master items for a category: ?category=workorder_type"""
    cat = (request.GET.get("category") or "").strip()
    from .models import MasterItem

    if not cat:
        return JsonResponse({"results": []})
    qs = MasterItem.objects.filter(category=cat, active=True).order_by("order", "label")
    out = []
    for m in qs:
        out.append({
            "id": m.id,
            "code": m.code or "",
            "label": m.label,
            "description": m.description or "",
            "meta": m.meta or {}
        })
    return JsonResponse({"results": out})


def api_equipment_list(request):
    """Return equipment list as JSON for AJAX calls (e.g., service request form)"""
    try:
        from datetime import date, timedelta
        
        # Get search query if provided
        q = request.GET.get("q", "").strip()
        pm_cal_due = request.GET.get("pm_cal_due", "").strip()
        
        # Query equipment from SQLite database
        equipments = Equipment_list.objects.all()
        
        # Filter for PM/CAL due equipment
        if pm_cal_due == "1":
            today = date.today()
            # Equipment due within next 30 days or overdue
            threshold = today + timedelta(days=30)
            from django.db.models import Q
            equipments = equipments.filter(
                Q(equipment_pm_due__lte=threshold) |
                Q(equipment_cal_due__lte=threshold)
            ).exclude(
                equipment_pm_due__isnull=True,
                equipment_cal_due__isnull=True
            )
        
        # Apply search filter if query provided
        if q:
            from django.db.models import Q
            equipments = equipments.filter(
                Q(equipment_id__icontains=q) |
                Q(equipment_code__icontains=q) |
                Q(equipment_name_EN__icontains=q) |
                Q(equipment_name_TH__icontains=q) |
                Q(equipment_user_customer__icontains=q)
            )
        
        # Limit results for performance - allow caller to request a larger limit via ?limit=
        # Default remains 100 to keep behavior stable; cap to avoid very large responses
        try:
            req_limit = int(request.GET.get('limit')) if request.GET.get('limit') else 100000
        except (TypeError, ValueError):
            req_limit = 100000
        # safety cap
        req_limit = min(max(req_limit, 1), 2000)
        equipments = equipments.order_by('equipment_id')[:req_limit]
        
        # Format response
        results = []
        for eq in equipments:
            # Provide explicit english/thai name fields and multiple key variants
            en_name = (getattr(eq, 'equipment_name_EN', '') or '')
            th_name = (getattr(eq, 'equipment_name_TH', '') or '')
            
            # Format due dates
            pm_due = getattr(eq, 'equipment_pm_due', None)
            cal_due = getattr(eq, 'equipment_cal_due', None)
            pm_due_str = pm_due.isoformat() if pm_due else ''
            cal_due_str = cal_due.isoformat() if cal_due else ''
            
            owner_val = getattr(eq, 'equipment_owner_customer', '') or ''
            results.append({
                    'code': eq.equipment_id or eq.equipment_code or '',
                    'name': th_name or en_name or '',
                    'equipment_name_EN': en_name,
                    'equipment_name_TH': th_name,
                    # also include alternate keys to be robust for client-side code
                    'name_EN': en_name,
                    'name_en': en_name,
                    'equipment_brand': eq.equipment_brand or '',
                    'equipment_model': eq.equipment_model or '',
                    'equipment_sn': eq.equipment_sn or '',
                    'equipment_gov': getattr(eq, 'equipment_gov', '') or '',
                    'equipment_user_customer': eq.equipment_user_customer or '',
                    # send canonical owner field and keep legacy alias for compatibility
                    'equipment_owner_customer': owner_val,
                    'equipment_owner': owner_val,
                    'equipment_image': getattr(eq, 'equipment_image', '') or '',
                    'equipment_code': eq.equipment_code or '',
                    'equipment_id': eq.equipment_id or '',
                    'equipment_pm_due': pm_due_str,
                    'equipment_cal_due': cal_due_str
                })
        
        return JsonResponse(results, safe=False)
    except Exception as e:
        logger.error(f"Error in api_equipment_list: {str(e)}")
        return JsonResponse({'error': str(e)}, status=500)


def create_request(request):
    """Proxy route: redirect to the Service Request creation page.

    Work orders must now originate from a Service Request,
    so direct users to create an SR first.
    """
    return redirect("service_request_create")


@require_http_methods(["GET", "POST"])
def service_request_create(request):
    """Public-facing endpoint for customers to submit a service request."""
    from .forms import ServiceRequestForm
    from .models import (ServiceRequestAttachment,
                         ServiceRequestLog)

    # Provide `profile` in context so templates can show user phone consistently
    profile = getattr(request.user, "profile", None) if request.user.is_authenticated else None

    # Provide master_customers (with ancestor chains) so templates can show parent info
    try:
        qs = list(MasterItem.objects.filter(category="customers", active=True).order_by("order", "label"))
        if qs:
            label_map = {m.id: m.label for m in qs}
            parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

            def build_ancestors(mid):
                chain = []
                seen = set()
                cur = parent_map.get(mid)
                while cur and cur not in seen:
                    seen.add(cur)
                    chain.insert(0, label_map.get(cur, ""))
                    cur = parent_map.get(cur)
                return " → ".join(chain) if chain else ""

            master_customers = [{"id": m.id, "label": m.label, "ancestors": build_ancestors(m.id)} for m in qs]
        else:
            master_customers = []
    except Exception:
        master_customers = []

    if request.method == "POST":
        form = ServiceRequestForm(request.POST, files=request.FILES)
        if form.is_valid():
            sr = form.save(commit=False)
            # if logged in, save as requested_by
            if request.user.is_authenticated:
                sr.requested_by = request.user
            # Persist preferred service date from Section 3 (field name: desired_service_date)
            try:
                desired_date_str = (request.POST.get("desired_service_date", "") or "").strip()
                if desired_date_str:
                    from datetime import datetime
                    sr.preferred_service_date = datetime.strptime(desired_date_str, "%Y-%m-%d").date()
            except Exception:
                # ignore invalid formats silently
                pass
            sr.save()
            # ── link pre-created appointment (auto-assign) if provided ──────────
            try:
                appt_id = (request.POST.get("appointment_id") or "").strip()
                if appt_id:
                    from .models import MaintenanceAppointment as _MA
                    _MA.objects.filter(id=int(appt_id)).update(notes=f"SR#{sr.id}")
            except Exception:
                pass
            # attachments
            if request.FILES.get("attachment"):
                f = request.FILES["attachment"]
                sra = ServiceRequestAttachment.objects.create(
                    request=sr,
                    uploaded_by=request.user if request.user.is_authenticated else None,
                )
                sra.file.save(f.name, f)
                sra.save()
            ServiceRequestLog.objects.create(
                request=sr,
                action="created",
                actor=request.user if request.user.is_authenticated else None,
                note="สร้างคำขอรับบริการ (Submitted)",
            )
            messages.success(request, "สร้างคำขอรับบริการเรียบร้อยแล้ว")
            # notify staff users (simple email) — best effort, fail silently
            try:
                from django.conf import settings
                from django.contrib.auth import get_user_model
                from django.core.mail import send_mail
                from django.template.loader import render_to_string

                User = get_user_model()
                staff_emails = list(
                    User.objects.filter(is_staff=True)
                    .exclude(email="")
                    .values_list("email", flat=True)
                )
                if staff_emails:
                    body = render_to_string(
                        "emails/sr_created.txt",
                        {
                            "sr": sr,
                            "site_url": getattr(
                                settings, "SITE_URL", "http://127.0.0.1:8000"
                            ),
                        },
                    )
                    send_mail(
                        subject=f"New Service Request #{sr.id}",
                        message=body,
                        from_email=None,
                        recipient_list=staff_emails,
                        fail_silently=True,
                    )
            except Exception:
                pass
            return redirect("my_service_requests")
    else:
        form = ServiceRequestForm()
    return render(request, "workorder/service_request_create.html", {"form": form, "profile": profile, "master_customers": master_customers})


def service_terms_page(request):
    """Public page rendering Service Terms content editable by admin via MasterItem(category='service_terms')."""
    try:
        from .models import MasterItem
        items = list(MasterItem.objects.filter(category="service_terms", active=True).order_by("order", "label"))
    except Exception:
        items = []

    # prefer first item as the primary content
    primary = items[0] if items else None
    return render(request, "service_terms.html", {"primary": primary, "items": items})


@user_passes_test(lambda u: u.is_superuser)
@require_http_methods(["GET", "POST"])
def service_terms_edit(request):
    """Superuser-only page to edit Service Terms (HTML/description) stored in MasterItem."""
    from .models import MasterItem

    # get or create primary item
    item = (
        MasterItem.objects.filter(category="service_terms").order_by("order", "label").first()
    )
    if not item:
        item = MasterItem.objects.create(
            category="service_terms",
            code="",
            label="Service Terms",
            description="",
            active=True,
            order=0,
            meta={}
        )

    if request.method == "POST":
        # Handle asset deletion
        del_url = (request.POST.get("delete_asset_url") or "").strip()
        if del_url:
            meta = item.meta or {}
            assets = list(meta.get("assets", []))
            # delete file from storage if under MEDIA_URL
            try:
                if del_url.startswith(settings.MEDIA_URL):
                    rel = del_url[len(settings.MEDIA_URL):]
                    default_storage.delete(rel)
            except Exception:
                pass
            # remove from list
            assets = [a for a in assets if a.get("url") != del_url]
            meta["assets"] = assets
            item.meta = meta
            item.save()
            messages.success(request, "ลบไฟล์ประกอบเรียบร้อย")
            return redirect("service_terms_edit")

        desc = (request.POST.get("description") or "").strip()
        html = (request.POST.get("content_html") or "").strip()
        item.description = desc
        meta = item.meta or {}
        meta["html"] = html
        # handle file uploads (assets)
        assets = meta.get("assets", [])
        upload_list = request.FILES.getlist("assets") or ([] if request.FILES.get("asset") is None else [request.FILES.get("asset")])
        for f in upload_list:
            # build unique path under service_terms/
            base, ext = os.path.splitext(f.name)
            unique_name = f"{timezone.now().strftime('%Y%m%d%H%M%S')}_{uuid.uuid4().hex}{ext}"
            path = os.path.join("service_terms", unique_name)
            saved_path = default_storage.save(path, f)
            url = settings.MEDIA_URL + saved_path.replace("\\", "/")
            assets.append({"name": f.name, "url": url})
        meta["assets"] = assets
        item.meta = meta
        item.save()
        messages.success(request, "บันทึกเงื่อนไขและไฟล์ประกอบเรียบร้อย")
        return redirect("service_terms_edit")

    return render(request, "service_terms_edit.html", {"item": item})


@permission_required("cmms.view_servicerequest", raise_exception=True)
def service_request_list(request):
    """Staff view: show all service requests (excluding deleted)"""
    from .models import ServiceRequest

    q = request.GET.get("q", "").strip()
    qs = ServiceRequest.objects.exclude(status="deleted")
    if q:
        qs = qs.filter(title__icontains=q)
    return render(
        request, "workorder/service_request_list.html", {"requests": qs, "q": q}
    )


@login_required
def my_service_requests(request):
    """User view: show only their own service requests"""
    from .models import ServiceRequest
    import re

    # Filter by current user
    if request.user.is_authenticated:
        qs = (
            ServiceRequest.objects
            .select_related("converted_to", "equipment")
            .filter(requested_by=request.user)
            .exclude(status="closed")  # hide cancelled/closed requests
            .order_by('-requested_at')
        )
    else:
        qs = ServiceRequest.objects.none()

    # Attach display fallbacks to each ServiceRequest instance so templates
    # can safely access `equipment_display_code` and `equipment_display_name`.
    try:
        from .models import Equipment_list
    except Exception:
        Equipment_list = None

    code_re = re.compile(r'รหัสเครื่อง[:\\s]*([A-Za-z0-9-]+)')
    name_re = re.compile(r'ชื่อเครื่อง[:\\s]*(.*?)(?=\\s*(แผนก|รหัสเครื่อง|กำหนด|$))')

    for r in qs:
        r.equipment_display_code = None
        r.equipment_display_name = None

        # Prefer linked equipment if present
        try:
            if getattr(r, 'equipment', None):
                eq = r.equipment
                r.equipment_display_code = getattr(eq, 'equipment_id', None)
                r.equipment_display_name = (
                    getattr(eq, 'equipment_name_EN', None)
                    or getattr(eq, 'equipment_name_TH', None)
                )
                continue
        except Exception:
            pass

        desc = (r.description or '')
        if desc:
            m = code_re.search(desc)
            if m:
                code = m.group(1).strip()
                r.equipment_display_code = code
                if Equipment_list is not None:
                    try:
                        eqobj = Equipment_list.objects.filter(equipment_id=code).first()
                        if eqobj:
                            r.equipment_display_name = (
                                getattr(eqobj, 'equipment_name_EN', None)
                                or getattr(eqobj, 'equipment_name_TH', None)
                            )
                    except Exception:
                        pass
            if not r.equipment_display_name:
                m2 = name_re.search(desc)
                if m2:
                    name = m2.group(1).strip()
                    if name:
                        r.equipment_display_name = name

    return render(
        request, "workorder/my_service_requests.html", {"requests": qs}
    )


@login_required
def service_request_cancel(request, sr_id):
    """Allow the request owner to cancel their Service Request before conversion.

    Rules:
    - Only the user who created the request can cancel it.
    - Allowed statuses to cancel: 'new', 'reviewed'.
    - If already converted, rejected, or closed, do nothing.
    """
    from .models import ServiceRequest, ServiceRequestLog
    from django.shortcuts import get_object_or_404

    sr = get_object_or_404(ServiceRequest, id=sr_id)
    # permission: only owner can cancel
    if sr.requested_by_id != getattr(request.user, "id", None):
        messages.error(request, "คุณไม่มีสิทธิ์ยกเลิกคำขอนี้")
        return redirect("my_service_requests")

    # must be POST for action
    if request.method != "POST":
        messages.error(request, "ไม่รองรับวิธีการเรียกใช้งานนี้")
        return redirect("my_service_requests")

    if sr.status not in ("new", "reviewed") or sr.converted_to_id:
        messages.warning(request, "คำขอถูกดำเนินการแล้ว ไม่สามารถยกเลิกได้")
        return redirect("my_service_requests")

    cancel_note = request.POST.get("note", "")
    # mark as closed to represent user-cancelled
    sr.status = "closed"
    if cancel_note:
        if sr.notes:
            sr.notes = sr.notes + "\n[Cancel note] " + cancel_note
        else:
            sr.notes = cancel_note
    sr.save()
    ServiceRequestLog.objects.create(
        request=sr,
        action="cancelled_by_requester",
        actor=request.user,
        note=cancel_note,
    )
    messages.success(request, "ยกเลิกคำขอเรียบร้อยแล้ว")
    return redirect("my_service_requests")


@login_required
def service_request_delete(request, sr_id):
    """Allow the request owner to delete their Service Request (hard delete).

    Rules:
    - Only the user who created the request can delete it.
    - Disallow deletion if the request was converted to a WorkOrder.
    - Must be POST.
    """
    from .models import ServiceRequest
    from django.shortcuts import get_object_or_404

    sr = get_object_or_404(ServiceRequest, id=sr_id)
    # permission: only owner can delete
    if sr.requested_by_id != getattr(request.user, "id", None):
        messages.error(request, "คุณไม่มีสิทธิ์ลบคำขอนี้")
        return redirect("my_service_requests")

    # Disallow deletion if converted
    if sr.converted_to_id:
        messages.warning(request, "คำขอถูกแปลงเป็นใบงานแล้ว ไม่สามารถลบได้")
        return redirect("my_service_requests")

    # GET -> render a small form asking for deletion reason
    if request.method == "GET":
        return render(
            request,
            "workorder/service_request_delete_reason.html",
            {"sr": sr},
        )

    # POST -> perform delete; require a reason
    reason = request.POST.get("reason", "").strip()
    if not reason:
        messages.error(request, "โปรดระบุสาเหตุการลบ")
        return redirect("service_request_delete", sr_id=sr.id)

    # create a log entry recording the deletion reason
    try:
        from .models import ServiceRequestLog

        ServiceRequestLog.objects.create(
            request=sr,
            action="deleted_by_requester",
            actor=request.user,
            note=reason,
        )
    except Exception:
        pass

    # Soft-delete: mark status as 'deleted' and preserve record in DB
    sr.status = "deleted"
    if reason:
        if sr.notes:
            sr.notes = sr.notes + "\n[Delete reason] " + reason
        else:
            sr.notes = "[Delete reason] " + reason
    sr.save()
    messages.success(request, "ย้ายคำขอไปยังรายการที่ถูกลบเรียบร้อยแล้ว")
    return redirect("my_service_requests")


@permission_required("cmms.view_servicerequest", raise_exception=True)
def service_request_detail(request, sr_id):
    from .models import (ServiceRequest, ServiceRequestLog)

    sr = ServiceRequest.objects.get(id=sr_id)
    if request.method == "POST":
        action = request.POST.get("action")
        # general save (notes, request_type)
        if (not action or action == "save") and request.user.has_perm(
            "cmms.change_servicerequest"
        ):
            # update notes and request_type
            sr.notes = request.POST.get("notes", "")
            rt = request.POST.get("request_type")
            if rt:
                sr.request_type = rt
            sr.save()
            try:
                ServiceRequestLog.objects.create(
                    request=sr,
                    action="updated",
                    actor=request.user,
                    note=f"Updated by {request.user.get_full_name() or request.user.username}",
                )
            except Exception:
                pass
            messages.success(request, "บันทึกการเปลี่ยนแปลงเรียบร้อยแล้ว")
            return redirect("service_request_detail", sr_id=sr.id)
        # officer can review and convert to WorkOrder
        if action == "mark_reviewed" and request.user.has_perm(
            "cmms.change_servicerequest"
        ):
            sr.status = "reviewed"
            sr.notes = request.POST.get("notes", "")
            sr.save()
            ServiceRequestLog.objects.create(
                request=sr, action="reviewed", actor=request.user, note=sr.notes
            )
            messages.success(request, "Marked as reviewed")
            return redirect("service_request_detail", sr_id=sr.id)

        if action == "convert_to_wo" and request.user.has_perm("cmms.add_workorder"):
            # collect optional convert note (from quick convert textarea)
            convert_note = request.POST.get(
                "convert_note", request.POST.get("notes", "")
            )
            # create WorkOrder from service request
            from cmms.models import WorkOrder

            wo = WorkOrder.objects.create(
                title=sr.title,
                description=sr.description,
                equipment=sr.equipment,
                reported_by=request.user,
            )
            sr.status = "converted"
            sr.converted_to = wo
            # save the convert note into sr.notes (append)
            if convert_note:
                if sr.notes:
                    sr.notes = sr.notes + "\n[Convert note] " + convert_note
                else:
                    sr.notes = convert_note
            sr.save()
            ServiceRequestLog.objects.create(
                request=sr,
                action="converted",
                actor=request.user,
                note=f"Converted to WO#{wo.id}. Note: {convert_note}",
            )
            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo,
                action="created_from_request",
                actor=request.user,
                note=f"From SR#{sr.id}",
            )
            messages.success(request, f"Converted to WorkOrder #{wo.id}")
            # notify requester (if email present) and staff
            try:
                from django.conf import settings
                from django.core.mail import send_mail
                from django.template.loader import render_to_string

                recipients = []
                if sr.requested_by and getattr(sr.requested_by, "email", ""):
                    recipients.append(sr.requested_by.email)
                if sr.customer_email:
                    recipients.append(sr.customer_email)
                from django.contrib.auth import get_user_model

                User = get_user_model()
                staff_emails = list(
                    User.objects.filter(is_staff=True)
                    .exclude(email="")
                    .values_list("email", flat=True)
                )
                for e in staff_emails:
                    if e not in recipients:
                        recipients.append(e)
                if recipients:
                    body = render_to_string(
                        "emails/sr_converted.txt",
                        {
                            "sr": sr,
                            "wo": wo,
                            "note": convert_note,
                            "site_url": getattr(
                                settings, "SITE_URL", "http://127.0.0.1:8000"
                            ),
                        },
                    )
                    send_mail(
                        subject=f"Service Request #{sr.id} Converted to WorkOrder #{wo.id}",
                        message=body,
                        from_email=None,
                        recipient_list=recipients,
                        fail_silently=True,
                    )
            except Exception:
                pass
            return redirect("workorder_detail", wo_id=wo.id)

    attachments = sr.attachments.all()
    logs = sr.logs.order_by("-created_at")

    # Use persisted preferred_service_date field (do not parse from notes)
    try:
        desired_service_date = getattr(sr, 'preferred_service_date', None)
    except Exception:
        desired_service_date = None

    return render(
        request,
        "workorder/service_request_detail.html",
        {
            "sr": sr,
            "attachments": attachments,
            "logs": logs,
            "desired_service_date": desired_service_date,
        },
    )


@permission_required("cmms.view_servicerequestlog", raise_exception=True)
def deleted_requests_report(request):
    """Show recent ServiceRequestLog entries where action == 'deleted_by_requester'."""
    from .models import ServiceRequestLog

    qs = (
        ServiceRequestLog.objects.filter(action="deleted_by_requester")
        .select_related("request", "actor")
        .order_by("-created_at")
    )

    return render(
        request,
        "workorder/service_requests_deleted.html",
        {"logs": qs},
    )


@permission_required("cmms.view_servicerequest", raise_exception=True)
def service_request_trash(request):
    """Show all deleted service requests (trash bin)."""
    from .models import ServiceRequest

    qs = ServiceRequest.objects.filter(status="deleted").order_by("-requested_at")
    return render(
        request,
        "workorder/service_request_trash.html",
        {"requests": qs},
    )


@login_required
def service_request_restore(request, sr_id):
    """Restore a deleted service request back to 'new' status.

    Rules:
    - Only staff with change_servicerequest permission or the original owner can restore.
    - Must be POST.
    """
    from .models import ServiceRequest, ServiceRequestLog
    from django.shortcuts import get_object_or_404

    sr = get_object_or_404(ServiceRequest, id=sr_id)

    # Check permission: staff or owner
    is_owner = sr.requested_by_id == getattr(request.user, "id", None)
    has_perm = request.user.has_perm("cmms.change_servicerequest")
    if not (is_owner or has_perm):
        messages.error(request, "คุณไม่มีสิทธิ์กู้คืนคำขอนี้")
        return redirect("service_request_trash")

    # Must be POST
    if request.method != "POST":
        messages.error(request, "ไม่รองรับวิธีการเรียกใช้งานนี้")
        return redirect("service_request_trash")

    # Only restore if currently deleted
    if sr.status != "deleted":
        messages.warning(request, "คำขอนี้ไม่ได้อยู่ในถังขยะ")
        return redirect("service_request_list")

    # Restore to 'new' status
    sr.status = "new"
    # Optionally append restore note
    restore_note = f"[Restored by {request.user.get_full_name() or request.user.username}]"
    if sr.notes:
        sr.notes = sr.notes + "\n" + restore_note
    else:
        sr.notes = restore_note
    sr.save()

    # Log the restore action
    ServiceRequestLog.objects.create(
        request=sr,
        action="restored",
        actor=request.user,
        note="Restored from trash",
    )

    messages.success(request, "กู้คืนคำขอเรียบร้อยแล้ว")
    return redirect("service_request_list")


@login_required
def service_request_purge(request, sr_id):
    """Permanently delete a ServiceRequest from the database (hard delete).

    Rules:
    - Only the original owner or users with 'cmms.delete_servicerequest' may purge.
    - Must be POST.
    """
    from .models import ServiceRequest, ServiceRequestLog
    from django.shortcuts import get_object_or_404

    sr = get_object_or_404(ServiceRequest, id=sr_id)

    # permission: owner or user with delete permission
    is_owner = sr.requested_by_id == getattr(request.user, "id", None)
    has_perm = request.user.has_perm("cmms.delete_servicerequest")
    if not (is_owner or has_perm):
        messages.error(request, "คุณไม่มีสิทธิ์ลบคำขอนี้ถาวร")
        return redirect("service_request_trash")

    if request.method != "POST":
        messages.error(request, "ไม่รองรับวิธีการเรียกใช้งานนี้")
        return redirect("service_request_trash")

    # create a log entry before deletion
    try:
        ServiceRequestLog.objects.create(
            request=sr,
            action="purged",
            actor=request.user,
            note="Permanently deleted from trash",
        )
    except Exception:
        pass

    sr.delete()
    messages.success(request, "ลบคำขอถาวรเรียบร้อยแล้ว")
    return redirect("service_request_trash")


def equipment_list(request):
    q = request.GET.get("q", "").strip()
    equipment_name_ENs_selected = request.GET.getlist("equipment_name_EN")
    equipment_name_THs_selected = request.GET.getlist("equipment_name_TH")
    equipment_brands_selected = request.GET.getlist("equipment_brand")
    equipment_models_selected = request.GET.getlist("equipment_model")
    # Accept both legacy param names (equipment_user_unit, etc.) and the
    # newer unified param `equipment_user_customer`. Combine them so both the
    # old templates and updated templates work without breaking filters.
    equipment_user_units_selected = (
        request.GET.getlist("equipment_user_unit") + request.GET.getlist("equipment_user_customer")
    )
    equipment_user_sections_selected = (
        request.GET.getlist("equipment_user_section") + request.GET.getlist("equipment_user_customer")
    )
    equipment_user_semi_departments_selected = (
        request.GET.getlist("equipment_user_semi_department") + request.GET.getlist("equipment_user_customer")
    )
    equipment_user_departments_selected = (
        request.GET.getlist("equipment_user_department") + request.GET.getlist("equipment_user_customer")
    )
    equipment_user_divisions_selected = (
        request.GET.getlist("equipment_user_division") + request.GET.getlist("equipment_user_customer")
    )
    equipment_user_customers_selected = request.GET.getlist("equipment_user_customer")

    all_equipment = Equipment_list.objects.all()
    if q:
        from django.db.models import Q

        # The model no longer contains the detailed user/unit/section fields.
        # Search only the remaining fields and the unified customer label.
        all_equipment = all_equipment.filter(
            Q(equipment_id__icontains=q)
            | Q(equipment_name_EN__icontains=q)
            | Q(equipment_name_TH__icontains=q)
            | Q(equipment_brand__icontains=q)
            | Q(equipment_model__icontains=q)
            | Q(equipment_sn__icontains=q)
            | Q(equipment_gov__icontains=q)
            | Q(equipment_owner_customer__icontains=q)
            | Q(equipment_user_customer__icontains=q)
        )
    if equipment_name_ENs_selected and any(
        val for val in equipment_name_ENs_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_name_EN__in=[val for val in equipment_name_ENs_selected if val]
        )
    if equipment_name_THs_selected and any(
        val for val in equipment_name_THs_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_name_TH__in=[val for val in equipment_name_THs_selected if val]
        )
    if equipment_brands_selected and any(
        val for val in equipment_brands_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_brand__in=[val for val in equipment_brands_selected if val]
        )
    if equipment_models_selected and any(
        val for val in equipment_models_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_model__in=[val for val in equipment_models_selected if val]
        )
    # Map any legacy user/unit/section filters to the unified equipment_user_customer
    if equipment_user_units_selected and any(val for val in equipment_user_units_selected if val):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[val for val in equipment_user_units_selected if val]
        )
    if equipment_user_sections_selected and any(val for val in equipment_user_sections_selected if val):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[val for val in equipment_user_sections_selected if val]
        )
    if equipment_user_semi_departments_selected and any(val for val in equipment_user_semi_departments_selected if val):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[val for val in equipment_user_semi_departments_selected if val]
        )
    if equipment_user_departments_selected and any(val for val in equipment_user_departments_selected if val):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[val for val in equipment_user_departments_selected if val]
        )
    if equipment_user_divisions_selected and any(val for val in equipment_user_divisions_selected if val):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[val for val in equipment_user_divisions_selected if val]
        )
    # Filter by customer (user/location)
    if equipment_user_customers_selected and any(
        val for val in equipment_user_customers_selected if val
    ):
        # Expand to include children of selected parents
        try:
            expanded_customers = []
            for selected_label in equipment_user_customers_selected:
                if selected_label:
                    expanded_customers.append(selected_label)
                    # Find children by searching for items that have this label as parent
                    parent_item = MasterItem.objects.filter(
                        label=selected_label, category='customers', active=True
                    ).first()
                    if parent_item:
                        children = MasterItem.objects.filter(
                            parent_id=parent_item.id, category='customers', active=True
                        )
                        for child in children:
                            expanded_customers.append(child.label)
            
            all_equipment = all_equipment.filter(
                equipment_user_customer__in=expanded_customers
            )
        except Exception:
            # Fallback to original filter if error
            all_equipment = all_equipment.filter(
                equipment_user_customer__in=[
                    val for val in equipment_user_customers_selected if val
                ]
            )
    
    # Filter by owner customer
    equipment_owner_customers_selected = request.GET.getlist("equipment_owner_customer")
    if equipment_owner_customers_selected and any(
        val for val in equipment_owner_customers_selected if val
    ):
        # Expand to include children of selected parents
        try:
            expanded_owners = []
            for selected_label in equipment_owner_customers_selected:
                if selected_label:
                    expanded_owners.append(selected_label)
                    # Find children by searching for items that have this label as parent
                    parent_item = MasterItem.objects.filter(
                        label=selected_label, category='customers', active=True
                    ).first()
                    if parent_item:
                        children = MasterItem.objects.filter(
                            parent_id=parent_item.id, category='customers', active=True
                        )
                        for child in children:
                            expanded_owners.append(child.label)
            
            all_equipment = all_equipment.filter(
                equipment_owner_customer__in=expanded_owners
            )
        except Exception:
            # Fallback to original filter if error
            all_equipment = all_equipment.filter(
                equipment_owner_customer__in=[
                    val for val in equipment_owner_customers_selected if val
                ]
            )
    
    # Filter by equipment status
    equipment_statuses_selected = request.GET.getlist("equipment_status")
    if equipment_statuses_selected and any(val for val in equipment_statuses_selected if val):
        # Assuming Equipment_list model has a status field (add if not exists)
        # For now, we'll pass the selection to template for client-side filtering
        # TODO: Add equipment_status field to Equipment_list model if needed
        pass

    # Helper function to get master items with ancestor chains
    def _ml(cat):
        """Return list of dicts: {id, label, ancestors: 'root > ... > parent'}"""
        qs = list(
            MasterItem.objects.filter(category=cat, active=True).order_by(
                "order", "label"
            )
        )
        if not qs:
            return []
        # build maps for fast lookup
        label_map = {m.id: m.label for m in qs}
        parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

        def build_ancestors(mid):
            chain = []
            seen = set()
            cur = parent_map.get(mid)
            while cur and cur not in seen:
                seen.add(cur)
                lbl = label_map.get(cur)
                if lbl:
                    chain.append(lbl)
                cur = parent_map.get(cur)
            chain.reverse()  # now root .. immediate_parent
            return " > ".join(chain)

        out = []
        for m in qs:
            out.append(
                {"id": m.id, "label": m.label, "ancestors": build_ancestors(m.id)}
            )
        return out
    
    # Helper function to expand parent selections to include children
    def expand_with_children(selected_labels, category):
        """Given parent labels, expand to include all children"""
        if not selected_labels:
            return []
        
        try:
            # Get all items in category
            all_items = MasterItem.objects.filter(category=category, active=True)
            label_to_id = {m.label: m.id for m in all_items}
            id_to_children = {}
            
            for item in all_items:
                if item.parent_id:
                    if item.parent_id not in id_to_children:
                        id_to_children[item.parent_id] = []
                    id_to_children[item.parent_id].append(item.label)
            
            # Expand selected labels
            expanded = set(selected_labels)
            for selected_label in selected_labels:
                selected_id = label_to_id.get(selected_label)
                if selected_id and selected_id in id_to_children:
                    expanded.update(id_to_children[selected_id])
            
            return list(expanded)
        except Exception as e:
            # If error occurs, just return original labels
            print(f"Error in expand_with_children: {e}")
            return selected_labels

    # สร้าง list สำหรับ dropdown
    equipment_name_ENs = (
        Equipment_list.objects.values_list("equipment_name_EN", flat=True)
        .distinct()
        .order_by("equipment_name_EN")
    )
    equipment_name_THs = (
        Equipment_list.objects.values_list("equipment_name_TH", flat=True)
        .distinct()
        .order_by("equipment_name_TH")
    )
    # Create hierarchical structure: brands -> models
    brand_model_data = (
        Equipment_list.objects.values("equipment_brand", "equipment_model")
        .distinct()
        .order_by("equipment_brand", "equipment_model")
    )
    
    # Build hierarchy
    equipment_brands_hierarchy = []
    brand_models_map = {}
    
    for item in brand_model_data:
        brand = item.get("equipment_brand") or ""
        model = item.get("equipment_model") or ""
        
        if brand and brand not in brand_models_map:
            brand_models_map[brand] = []
            equipment_brands_hierarchy.append({
                "label": brand,
                "ancestors": "",
                "is_parent": True
            })
        
        if brand and model:
            brand_models_map[brand].append(model)
    
    # Add models as children of brands
    equipment_models_hierarchy = []
    for brand, models in brand_models_map.items():
        for model in models:
            equipment_models_hierarchy.append({
                "label": model,
                "ancestors": brand,
                "parent": brand
            })
    
    # Keep original flat lists for backward compatibility
    equipment_brands = list(brand_models_map.keys())
    equipment_models = [m["label"] for m in equipment_models_hierarchy]
    # Use MasterItem category 'customers' for units (site / owning unit) as requested
    try:
        # prefer master items with full ancestor chains
        equipment_user_units = _ml("customers")
    except Exception:
        equipment_user_units = list(
            Equipment_list.objects.values_list("equipment_user_customer", flat=True)
            .distinct()
            .order_by("equipment_user_customer")
        )
    equipment_user_sections = (
        Equipment_list.objects.values_list("equipment_user_customer", flat=True)
        .distinct()
        .order_by("equipment_user_customer")
    )
    equipment_user_semi_departments = (
        Equipment_list.objects.values_list("equipment_user_customer", flat=True)
        .distinct()
        .order_by("equipment_user_customer")
    )
    equipment_user_departments = (
        Equipment_list.objects.values_list("equipment_user_customer", flat=True)
        .distinct()
        .order_by("equipment_user_customer")
    )
    equipment_user_divisions = (
        Equipment_list.objects.values_list("equipment_user_customer", flat=True)
        .distinct()
        .order_by("equipment_user_customer")
    )
    # For "customer" (higher level), also prefer MasterItem category 'customers' if available
    try:
        equipment_user_customers = _ml("customers")
    except Exception:
        equipment_user_customers = list(
            Equipment_list.objects.values_list("equipment_user_customer", flat=True)
            .distinct()
            .order_by("equipment_user_customer")
        )

    # Get owner customers (same as user customers - using MasterItem 'customers')
    try:
        equipment_owner_customers = _ml("customers")
    except Exception:
        equipment_owner_customers = list(
            Equipment_list.objects.values_list("equipment_owner_customer", flat=True)
            .distinct()
            .order_by("equipment_owner_customer")
        )
    
    # Get selected owner customers from request
    equipment_owner_customers_selected = request.GET.getlist("equipment_owner_customer")

    # Get master status items for dynamic status rendering and filtering
    master_statuses = MasterItem.objects.filter(
        category='equipment_status', 
        active=True
    ).order_by('order', 'label')
    
    # Get selected status labels from request
    equipment_statuses_selected = request.GET.getlist("equipment_status")
    
    # Helper function for building equipment master list
    def _ml_equipments():
        # Build lookup dictionaries for value -> label mapping directly from MasterItem
        # (Don't use _ml() here because it doesn't return the 'code' field)
        equipment_type_items = MasterItem.objects.filter(category='equipment_type', active=True)
        lookup_type = {}
        for item in equipment_type_items:
            if item.code:
                lookup_type[item.code] = item.label
            lookup_type[item.label] = item.label
        
        # Also build lookup for PM/CAL frequencies
        pm_fq_items = MasterItem.objects.filter(category='frequency_pm', active=True)
        lookup_pm_fq = {}
        for item in pm_fq_items:
            if item.code:
                lookup_pm_fq[item.code] = item.label
            lookup_pm_fq[item.label] = item.label
        
        cal_fq_items = MasterItem.objects.filter(category='frequency_cal', active=True)
        lookup_cal_fq = {}
        for item in cal_fq_items:
            if item.code:
                lookup_cal_fq[item.code] = item.label
            lookup_cal_fq[item.label] = item.label
        
        # Also build lookup for Risk
        risk_items = MasterItem.objects.filter(category='risks', active=True)
        lookup_risk = {}
        for item in risk_items:
            if item.code:
                lookup_risk[item.code] = item.label
            lookup_risk[item.label] = item.label
        
        qs = list(
            MasterItem.objects.filter(category='equipments', active=True).order_by(
                "order", "code"
            )
        )
        out = []
        for m in qs:
            meta = m.meta or {}
            type_val = meta.get("equipment_type", "")
            lifespan_val = meta.get("equipment_life", "") or meta.get("lifespan", "")
            type_label = lookup_type.get(type_val, "")
            # Accept multiple possible meta key names for PM/CAL frequency
            pm_fq_val = (
                meta.get("pm_fq")
                or meta.get("equipment_pm_fq")
                or meta.get("frequency_pm")
                or ""
            )
            cal_fq_val = (
                meta.get("cal_fq")
                or meta.get("equipment_cal_fq")
                or meta.get("frequency_cal")
                or ""
            )
            risk_val = meta.get("equipment_risk", "")
            pm_fq_label = lookup_pm_fq.get(pm_fq_val, "")
            cal_fq_label = lookup_cal_fq.get(cal_fq_val, "")
            risk_label = lookup_risk.get(risk_val, "")
            
            # Save equipment_type_label to meta for persistence
            if type_label and type_label != meta.get("equipment_type_label"):
                meta["equipment_type_label"] = type_label
                m.meta = meta
                m.save()
            
            out.append({
                "id": m.id,
                "code": m.code,
                "label": m.label,
                "name_th": meta.get("equipment_name_th", ""),
                "equipment_type": type_val,
                "equipment_type_label": type_label,
                "lifespan": lifespan_val,
                "pm_fq": pm_fq_val,
                "pm_fq_label": pm_fq_label,
                "cal_fq": cal_fq_val,
                "cal_fq_label": cal_fq_label,
                "risk": risk_val,
                "risk_label": risk_label,
            })
        return out
    
    return render(
        request,
        "equipment/equipment_list.html",
        {
            "all_equipment": all_equipment,
            "equipment_name_ENs": equipment_name_ENs,
            "equipment_name_THs": equipment_name_THs,
            "equipment_brands": equipment_brands,
            "equipment_models": equipment_models,
            "equipment_brands_hierarchy": equipment_brands_hierarchy,
            "equipment_models_hierarchy": equipment_models_hierarchy,
            "equipment_user_units": equipment_user_units,
            "equipment_user_sections": equipment_user_sections,
            "equipment_user_semi_departments": equipment_user_semi_departments,
            "equipment_user_departments": equipment_user_departments,
            "equipment_user_divisions": equipment_user_divisions,
            "equipment_user_customers": equipment_user_customers,
            "equipment_owner_customers": equipment_owner_customers,
            "equipment_name_ENs_selected": equipment_name_ENs_selected,
            "equipment_name_THs_selected": equipment_name_THs_selected,
            "equipment_brands_selected": equipment_brands_selected,
            "equipment_models_selected": equipment_models_selected,
            "equipment_user_units_selected": equipment_user_units_selected,
            "equipment_user_sections_selected": equipment_user_sections_selected,
            "equipment_user_departments_selected": equipment_user_departments_selected,
            "equipment_user_divisions_selected": equipment_user_divisions_selected,
            "equipment_user_customers_selected": equipment_user_customers_selected,
            "equipment_owner_customers_selected": equipment_owner_customers_selected,
            "equipment_statuses_selected": equipment_statuses_selected,
            "master_statuses": master_statuses,
            "master_equipments": _ml_equipments(),
            "master_frequency_pm": _ml("frequency_pm"),
            "master_frequency_cal": _ml("frequency_cal"),
        },
    )


def _validate_equipment_image(upload):
    if upload is None:
        return None
    allowed_types = {
        ".jpg": ("image/jpeg", lambda header: header.startswith(b"\xff\xd8\xff")),
        ".jpeg": ("image/jpeg", lambda header: header.startswith(b"\xff\xd8\xff")),
        ".png": ("image/png", lambda header: header.startswith(b"\x89PNG\r\n\x1a\n")),
        ".gif": ("image/gif", lambda header: header.startswith((b"GIF87a", b"GIF89a"))),
        ".webp": (
            "image/webp",
            lambda header: header.startswith(b"RIFF") and header[8:12] == b"WEBP",
        ),
        ".bmp": ("image/bmp", lambda header: header.startswith(b"BM")),
    }
    extension = os.path.splitext(upload.name)[1].lower()
    if extension not in allowed_types:
        return "กรุณาเลือกไฟล์ภาพ JPG, PNG, GIF, WEBP หรือ BMP"
    expected_type, signature_check = allowed_types[extension]
    header = upload.read(12)
    upload.seek(0)
    if upload.content_type != expected_type or not signature_check(header):
        return "ไฟล์ที่เลือกไม่ใช่รูปภาพตามชนิดไฟล์ที่กำหนด"
    if upload.size > 10 * 1024 * 1024:
        return "ไฟล์ภาพต้องมีขนาดไม่เกิน 10 MB"
    return None


def _save_equipment_image(equipment, upload):
    from .models import equipment_image_upload_path

    old_name = equipment.equipment_image.name
    target_name = equipment_image_upload_path(equipment, upload.name)
    if default_storage.exists(target_name):
        default_storage.delete(target_name)
    equipment.equipment_image.save(os.path.basename(target_name), upload, save=True)
    if old_name and old_name != equipment.equipment_image.name:
        default_storage.delete(old_name)


def _equipment_image_exists(equipment):
    image = equipment.equipment_image
    return bool(image and image.storage.exists(image.name))


@login_required
def equipment_images(request):
    equipment_with_images = (
        Equipment_list.objects.exclude(equipment_image="")
        .order_by("equipment_id")
    )
    equipment = [
        {
            "equipment": item,
            "image_available": _equipment_image_exists(item),
        }
        for item in equipment_with_images
    ]
    return render(
        request,
        "equipment/equipment_images.html",
        {"equipment_images": equipment},
    )


@login_required
def equipment_image_view(request, equipment_list_id):
    equipment = get_object_or_404(Equipment_list, id=equipment_list_id)
    if not _equipment_image_exists(equipment):
        return HttpResponse(status=404)
    image = equipment.equipment_image
    content_type = mimetypes.guess_type(image.name)[0] or "application/octet-stream"
    return FileResponse(
        image.open("rb"),
        content_type=content_type,
        as_attachment=False,
        filename=os.path.basename(image.name),
    )


@login_required
def equipment_image_download(request, equipment_list_id):
    equipment = get_object_or_404(Equipment_list, id=equipment_list_id)
    if not _equipment_image_exists(equipment):
        return HttpResponse(status=404)
    image = equipment.equipment_image
    content_type = mimetypes.guess_type(image.name)[0] or "application/octet-stream"
    return FileResponse(
        image.open("rb"),
        content_type=content_type,
        as_attachment=True,
        filename=os.path.basename(image.name),
    )


@permission_required("cmms.add_equipment_list", raise_exception=True)
def add_equipment(request):
    if request.method == "POST":
        # Required fields expected by the front-end form. Keep this list
        # reasonably small because the DB model is permissive; validation
        # for visible required fields happens in the client UI.
        required_fields = [
            "equipment_id",
            "equipment_code",
            "equipment_name_EN",
            "equipment_name_TH",
            "equipment_brand",
            "equipment_model",
            "equipment_sn",
            "equipment_gov",
            "equipment_price",
            "equipment_photo",
            "equipment_type",
            "equipment_life",
            "equipment_waranty_date",
            "equipment_waranty_due",
            "equipment_distributor_name",
            "equipment_distributor_tel",
            "equipment_pm_fq",
            "equipment_pm_due",
            "equipment_cal_fq",
            "equipment_cal_due",
            "equipment_owner_customer",
            "equipment_user_customer",
            "equipment_register_username",
            "equipment_register_adminname",
            "equipment_note",
        ]

        numeric_fields = [
            "equipment_price",
            "equipment_pm_fq",
            "equipment_cal_fq",
            "equipment_life",
        ]
        date_fields = [
            "equipment_waranty_date",
            "equipment_waranty_due",
            "equipment_pm_due",
            "equipment_cal_due",
        ]

        form_values = {}
        errors = {}
        image_upload = request.FILES.get("equipment_image")
        # Collect values (no longer require presence for any field)
        for field in required_fields:
            val = request.POST.get(field, "")
            if isinstance(val, str):
                val = val.strip()
            form_values[field] = val
            # Removed: required validation (allow empty values)

        image_error = _validate_equipment_image(image_upload)
        if image_error:
            errors["equipment_image"] = image_error

        # numeric validation
        for field in numeric_fields:
            v = form_values.get(field, "")
            if v != "":
                try:
                    int(v)
                except Exception:
                    errors[field] = "ต้องเป็นตัวเลข"

        # equipment_id validation (only when provided)
        equipment_id_val = form_values.get("equipment_id", "")
        if equipment_id_val != "":
            # minimum length check
            if len(equipment_id_val) < 13:
                errors["equipment_id"] = "รหัสต้องมีอย่างน้อย 13 ตัวอักษร"
            # uniqueness check (prevent duplicate equipment_id on create)
            elif Equipment_list.objects.filter(equipment_id=equipment_id_val).exists():
                errors["equipment_id"] = "รหัสนี้มีอยู่แล้ว"

        # date validation (expecting YYYY-MM-DD)
        import datetime

        for field in date_fields:
            v = form_values.get(field, "")
            if v != "":
                try:
                    datetime.date.fromisoformat(v)
                except Exception:
                    errors[field] = "รูปแบบวันที่ไม่ถูกต้อง (YYYY-MM-DD)"

        if errors:
            # prepare master lists (same shape as GET) so the unified template can render selects
            def _ml_err(cat):
                qs = list(
                    MasterItem.objects.filter(category=cat, active=True).order_by(
                        "order", "label"
                    )
                )
                if not qs:
                    return []
                label_map = {m.id: m.label for m in qs}
                parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

                def build_ancestors(mid):
                    chain = []
                    seen = set()
                    cur = parent_map.get(mid)
                    while cur and cur not in seen:
                        seen.add(cur)
                        lbl = label_map.get(cur)
                        if lbl:
                            chain.append(lbl)
                        cur = parent_map.get(cur)
                    chain.reverse()
                    return " > ".join(chain)

                out = []
                for m in qs:
                    out.append(
                        {
                            "id": m.id,
                            "label": m.label,
                            "ancestors": build_ancestors(m.id),
                        }
                    )
                return out

            context = {
                "errors": errors,
                "form_values": form_values,
                "master_equipment_type": _ml_err("equipment_type"),
                "master_manufacturers": _ml_err("manufacturers"),
                "master_departments": _ml_err("departments"),
                "master_units": _ml_err("units"),
                "master_customers": _ml_err("customers"),
                "master_locations": _ml_err("locations"),
                "master_service_providers": _ml_err("service_providers"),
            }
            return render(request, "equipment/equipment_form.html", context)

        # All validations passed — create the object with proper types
        # Cast numeric and date fields
        def to_int(name):
            return int(form_values.get(name)) if form_values.get(name) != "" else None

        def to_date(name):
            return (
                datetime.date.fromisoformat(form_values.get(name))
                if form_values.get(name) != ""
                else None
            )

        equipment_list = Equipment_list.objects.create(
            equipment_id=form_values.get("equipment_id") or "",
            equipment_code=form_values.get("equipment_code") or "",
            equipment_name_EN=form_values.get("equipment_name_EN") or "",
            equipment_name_TH=form_values.get("equipment_name_TH") or "",
            equipment_brand=form_values.get("equipment_brand") or "",
            equipment_model=form_values.get("equipment_model") or "",
            equipment_sn=form_values.get("equipment_sn") or "",
            equipment_gov=form_values.get("equipment_gov") or "",
            equipment_price=to_int("equipment_price"),
            equipment_photo=form_values.get("equipment_photo") or "",
            equipment_type=form_values.get("equipment_type") or "",
            equipment_life=to_int("equipment_life"),
            equipment_waranty_date=to_date("equipment_waranty_date"),
            equipment_waranty_due=to_date("equipment_waranty_due"),
            equipment_distributor_name=form_values.get("equipment_distributor_name") or "",
            equipment_distributor_tel=form_values.get("equipment_distributor_tel") or "",
            equipment_pm_fq=to_int("equipment_pm_fq"),
            equipment_pm_due=to_date("equipment_pm_due"),
            equipment_cal_fq=to_int("equipment_cal_fq"),
            equipment_cal_due=to_date("equipment_cal_due"),
            equipment_owner_customer=form_values.get("equipment_owner_customer") or "",
            equipment_user_customer=form_values.get("equipment_user_customer") or "",
            equipment_service_provider=form_values.get("equipment_service_provider") or "",
            equipment_register_username=form_values.get("equipment_register_username") or "",
            equipment_register_adminname=form_values.get(
                "equipment_register_adminname"
            ) or "",
            equipment_note=form_values.get("equipment_note") or "",
            created_by=request.user.get_full_name() or request.user.username if request.user.is_authenticated else '',
        )
        equipment_list.save()
        if image_upload:
            _save_equipment_image(equipment_list, image_upload)
        
        # Log creation history
        new_data = capture_equipment_snapshot(equipment_list)
        log_equipment_history(
            equipment_id=equipment_list.equipment_id,
            action_type='CREATE',
            user=request.user,
            request=request,
            new_data=new_data,
            notes=f'สร้างทะเบียนอุปกรณ์ใหม่'
        )
        
        messages.success(request, "บันทึกข้อมูลเรียบร้อยแล้ว")
        return redirect("/equipment_list")
    else:
        # show latest equipment_id as default in the form (if any)
        # Use the last value in the equipment_id column (ordered by equipment_id)
        latest_equipment = Equipment_list.objects.order_by("equipment_id").last()
        latest_code = latest_equipment.equipment_id if latest_equipment else "MEDCMU-"

        # build a suggested next code by incrementing trailing number if present
        suggested_code = latest_code
        import re

        m = re.search(r"^(.*?)(\d+)$", latest_code)
        if m:
            prefix = m.group(1)
            num = m.group(2)
            try:
                next_num = str(int(num) + 1).zfill(len(num))
                suggested_code = f"{prefix}{next_num}"
            except Exception:
                suggested_code = latest_code
        else:
            # no trailing number -> append 1
            suggested_code = latest_code + "1"

    # gather master lists for selects
    def _ml(cat):
        # return list of dicts: {label: ..., ancestors: 'root > ... > parent'}
        qs = list(
            MasterItem.objects.filter(category=cat, active=True).order_by(
                "order", "label"
            )
        )
        if not qs:
            return []
        # build maps for fast lookup
        label_map = {m.id: m.label for m in qs}
        parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

        def build_ancestors(mid):
            chain = []
            seen = set()
            cur = parent_map.get(mid)
            while cur and cur not in seen:
                seen.add(cur)
                lbl = label_map.get(cur)
                if lbl:
                    chain.append(lbl)
                cur = parent_map.get(cur)
            # chain currently [immediate_parent, parent_of_parent, ...]
            chain.reverse()  # now root .. immediate_parent
            return " > ".join(chain)

        out = []
        for m in qs:
            out.append(
                {"id": m.id, "code": m.code, "label": m.label, "ancestors": build_ancestors(m.id)}
            )
        return out

    # Build master_equipments list with code, label, and meta fields
    def _ml_equipments():
        # Build lookup dictionaries for value -> label mapping
        # Use both code and label as keys (some master data use label as value)
        def build_lookup(items):
            d = {}
            for item in items:
                if item.get('code'):
                    d[item['code']] = item['label']
                d[item['label']] = item['label']  # label -> label mapping
            return d
        
        lookup_risk = build_lookup(_ml("risks"))
        lookup_type = build_lookup(_ml("equipment_type"))
        lookup_pm = build_lookup(_ml("frequency_pm"))
        lookup_cal = build_lookup(_ml("frequency_cal"))
        lookup_sv_provider = build_lookup(_ml("service_providers"))
        
        qs = list(
            MasterItem.objects.filter(category='equipments', active=True).order_by(
                "order", "code"
            )
        )
        out = []
        for m in qs:
            meta = m.meta or {}
            # Accept several possible meta key names for PM/CAL frequency because
            # imported master data has used different keys historically (e.g. "frequency_pm", "frequency_cal").
            pm_val = (
                meta.get("equipment_pm_fq")
                or meta.get("pm_fq")
                or meta.get("frequency_pm")
                or meta.get("frequency")
                or ""
            )
            cal_val = (
                meta.get("equipment_cal_fq")
                or meta.get("cal_fq")
                or meta.get("frequency_cal")
                or meta.get("frequency")
                or ""
            )
            risk_val = meta.get("equipment_risk", "")
            type_val = meta.get("equipment_type", "")
            lifespan_val = meta.get("equipment_life", "") or meta.get("lifespan", "")
            sv_provider_val = meta.get("equipment_service_provider", "") or meta.get("sv_provider", "")
            
            # Convert values to labels using lookup dictionaries
            risk_label = lookup_risk.get(risk_val, "")
            pm_fq_label = lookup_pm.get(pm_val, "")
            cal_fq_label = lookup_cal.get(cal_val, "")
            type_label = lookup_type.get(type_val, "")
            sv_provider_label = lookup_sv_provider.get(sv_provider_val, "")
            
            out.append({
                "id": m.id,
                "code": m.code,
                "label": m.label,
                "name_en": meta.get("equipment_name_en", ""),
                "name_th": meta.get("equipment_name_th", ""),
                # Additional master metadata used in equipment form/modal
                "risk": risk_val,
                "risk_label": risk_label,
                "pm_fq": pm_val,
                "pm_fq_label": pm_fq_label,
                "cal_fq": cal_val,
                "cal_fq_label": cal_fq_label,
                "equipment_type": type_val,
                "equipment_type_label": type_label,
                "lifespan": lifespan_val,
                "sv_provider": sv_provider_val,
                "sv_provider_label": sv_provider_label,
            })
        return out

    context = {
        "form_values": {
            "equipment_id": latest_code,
            "suggested_equipment_id": suggested_code,
        },
        "master_equipment_type": _ml("equipment_type"),
        "master_frequency_cal": _ml("frequency_cal"),
        "master_frequency_pm": _ml("frequency_pm"),
        "master_manufacturers": _ml("manufacturers"),
        "master_departments": _ml("departments"),
        "master_units": _ml("units"),
        "master_customers": _ml("customers"),
        "master_locations": _ml("locations"),
        "master_service_providers": _ml("service_providers"),
        "master_equipments": _ml_equipments(),
    }
    return render(request, "equipment/equipment_form.html", context)


@permission_required("cmms.change_equipment_list", raise_exception=True)
def edit_equipment(request, equipment_list_id):
    if request.method == "POST":
        equipment_list = Equipment_list.objects.get(id=equipment_list_id)
        image_upload = request.FILES.get("equipment_image")
        image_error = _validate_equipment_image(image_upload)
        if image_error:
            messages.error(request, image_error)
            return redirect("edit_equipment", equipment_list_id=equipment_list_id)
        # prevent assigning a duplicate equipment_id when editing
        new_equipment_id = request.POST.get("equipment_id", "").strip()
        if (
            new_equipment_id
            and Equipment_list.objects.filter(equipment_id=new_equipment_id)
            .exclude(id=equipment_list_id)
            .exists()
        ):
            messages.error(request, "รหัสนี้มีอยู่แล้ว ไม่สามารถแก้ไขเป็นรหัสนี้ได้")

            # build form_values from existing object so unified template can render
            def _fv_from_obj(obj):
                def _fmt(v):
                    try:
                        return (
                            v.isoformat()
                            if hasattr(v, "isoformat")
                            else (v if v is not None else "")
                        )
                    except Exception:
                        return v

                return {
                    "equipment_id": getattr(obj, "equipment_id", ""),
                    "equipment_code": getattr(obj, "equipment_code", ""),
                    "equipment_name_EN": getattr(obj, "equipment_name_EN", ""),
                    "equipment_name_TH": getattr(obj, "equipment_name_TH", ""),
                    "equipment_brand": getattr(obj, "equipment_brand", ""),
                    "equipment_model": getattr(obj, "equipment_model", ""),
                    "equipment_sn": getattr(obj, "equipment_sn", ""),
                    "equipment_gov": getattr(obj, "equipment_gov", ""),
                    "equipment_price": getattr(obj, "equipment_price", ""),
                    "equipment_photo": getattr(obj, "equipment_photo", ""),
                    "equipment_type": getattr(obj, "equipment_type", ""),
                    "equipment_life": getattr(obj, "equipment_life", ""),
                    "equipment_waranty_date": _fmt(
                        getattr(obj, "equipment_waranty_date", "")
                    ),
                    "equipment_waranty_due": _fmt(
                        getattr(obj, "equipment_waranty_due", "")
                    ),
                    "equipment_distributor_name": getattr(
                        obj, "equipment_distributor_name", ""
                    ),
                    "equipment_distributor_tel": getattr(
                        obj, "equipment_distributor_tel", ""
                    ),
                    "equipment_pm_fq": getattr(obj, "equipment_pm_fq", ""),
                    "equipment_pm_due": _fmt(getattr(obj, "equipment_pm_due", "")),
                    "equipment_cal_fq": getattr(obj, "equipment_cal_fq", ""),
                    "equipment_cal_due": _fmt(getattr(obj, "equipment_cal_due", "")),
                    "equipment_owner_customer": getattr(
                        obj, "equipment_owner_customer", ""
                    ),
                    "equipment_user_customer": getattr(
                        obj, "equipment_user_customer", ""
                    ),
                    "equipment_service_provider": getattr(
                        obj, "equipment_service_provider", ""
                    ),
                    "equipment_register_username": getattr(
                        obj, "equipment_register_username", ""
                    ),
                    "equipment_register_adminname": getattr(
                        obj, "equipment_register_adminname", ""
                    ),
                    "equipment_note": getattr(obj, "equipment_note", ""),
                }

            form_values = _fv_from_obj(equipment_list)

            # prepare master lists with ancestors (as add form expects)
            def _ml(cat):
                qs = list(
                    MasterItem.objects.filter(category=cat, active=True).order_by(
                        "order", "label"
                    )
                )
                if not qs:
                    return []
                label_map = {m.id: m.label for m in qs}
                parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

                def build_ancestors(mid):
                    chain = []
                    seen = set()
                    cur = parent_map.get(mid)
                    while cur and cur not in seen:
                        seen.add(cur)
                        lbl = label_map.get(cur)
                        if lbl:
                            chain.append(lbl)
                        cur = parent_map.get(cur)
                    chain.reverse()
                    return " > ".join(chain)

                out = []
                for m in qs:
                    out.append(
                        {
                            "id": m.id,
                            "label": m.label,
                            "ancestors": build_ancestors(m.id),
                        }
                    )
                return out

            context = {
                "form_values": form_values,
                "edit": True,
                "equipment_list": equipment_list,
                "has_equipment_image": _equipment_image_exists(equipment_list),
                "master_equipment_type": _ml("equipment_type"),
                "master_frequency_cal": _ml("frequency_cal"),
                "master_frequency_pm": _ml("frequency_pm"),
                "master_manufacturers": _ml("manufacturers"),
                "master_departments": _ml("departments"),
                "master_units": _ml("units"),
                "master_customers": _ml("customers"),
                "master_locations": _ml("locations"),
                "master_service_providers": _ml("service_providers"),
            }
            return render(request, "equipment/equipment_form.html", context)
        # Capture old data before updating for history tracking
        old_data = capture_equipment_snapshot(equipment_list)
        
        # Import datetime for date field handling
        import datetime
        
        # Helper functions to convert empty strings to None for nullable fields
        def to_int_or_none(value):
            """Convert to int or None. Returns None for empty strings."""
            if value == "" or value is None:
                return None
            try:
                return int(value)
            except (ValueError, TypeError):
                return None
        
        def to_date_or_none(value):
            """Convert to date or None. Returns None for empty strings."""
            if value == "" or value is None:
                return None
            try:
                return datetime.date.fromisoformat(value)
            except (ValueError, TypeError):
                return None
        
        # Only set fields that exist on the current model. Older code
        # attempted to set removed fields (unit/section/division, pm_date,
        # cal_date, etc.) which would raise AttributeError after the schema
        # change.
        equipment_list.equipment_id = request.POST.get("equipment_id", equipment_list.equipment_id)
        equipment_list.equipment_code = request.POST.get("equipment_code", equipment_list.equipment_code)
        equipment_list.equipment_name_EN = request.POST.get("equipment_name_EN", equipment_list.equipment_name_EN)
        equipment_list.equipment_name_TH = request.POST.get("equipment_name_TH", equipment_list.equipment_name_TH)
        equipment_list.equipment_brand = request.POST.get("equipment_brand", equipment_list.equipment_brand)
        equipment_list.equipment_model = request.POST.get("equipment_model", equipment_list.equipment_model)
        equipment_list.equipment_sn = request.POST.get("equipment_sn", equipment_list.equipment_sn)
        equipment_list.equipment_gov = request.POST.get("equipment_gov", equipment_list.equipment_gov)
        equipment_list.equipment_price = to_int_or_none(request.POST.get("equipment_price", ""))
        equipment_list.equipment_photo = request.POST.get("equipment_photo", equipment_list.equipment_photo)
        equipment_list.equipment_type = request.POST.get("equipment_type", equipment_list.equipment_type)
        equipment_list.equipment_life = to_int_or_none(request.POST.get("equipment_life", ""))
        equipment_list.equipment_waranty_date = to_date_or_none(request.POST.get("equipment_waranty_date", ""))
        equipment_list.equipment_waranty_due = to_date_or_none(request.POST.get("equipment_waranty_due", ""))
        equipment_list.equipment_distributor_name = request.POST.get("equipment_distributor_name", equipment_list.equipment_distributor_name)
        equipment_list.equipment_distributor_tel = request.POST.get("equipment_distributor_tel", equipment_list.equipment_distributor_tel)
        equipment_list.equipment_pm_fq = to_int_or_none(request.POST.get("equipment_pm_fq", ""))
        equipment_list.equipment_pm_due = to_date_or_none(request.POST.get("equipment_pm_due", ""))
        equipment_list.equipment_cal_fq = to_int_or_none(request.POST.get("equipment_cal_fq", ""))
        equipment_list.equipment_cal_due = to_date_or_none(request.POST.get("equipment_cal_due", ""))
        equipment_list.equipment_owner_customer = request.POST.get("equipment_owner_customer", equipment_list.equipment_owner_customer)
        equipment_list.equipment_user_customer = request.POST.get("equipment_user_customer", equipment_list.equipment_user_customer)
        equipment_list.equipment_service_provider = request.POST.get("equipment_service_provider", equipment_list.equipment_service_provider)
        equipment_list.equipment_register_username = request.POST.get("equipment_register_username", equipment_list.equipment_register_username)
        equipment_list.equipment_register_adminname = request.POST.get("equipment_register_adminname", equipment_list.equipment_register_adminname)
        equipment_list.equipment_note = request.POST.get("equipment_note", equipment_list.equipment_note)
        
        # Set updated_by before saving
        equipment_list.updated_by = request.user.get_full_name() or request.user.username if request.user.is_authenticated else ''
        
        if image_upload:
            _save_equipment_image(equipment_list, image_upload)
        else:
            equipment_list.save()
        
        # Log edit history
        new_data = capture_equipment_snapshot(equipment_list)
        log_equipment_history(
            equipment_id=equipment_list.equipment_id,
            action_type='UPDATE',
            user=request.user,
            request=request,
            old_data=old_data,
            new_data=new_data,
            notes=f'แก้ไขทะเบียนอุปกรณ์'
        )
        
        messages.success(request, "แก้ไขข้อมูลเรียบร้อยแล้ว")
        return redirect("/equipment_list")

    else:
        # ดึงข้อมูลเก่าขึ้นมาแสดง
        equipment_list = Equipment_list.objects.get(id=equipment_list_id)

    # build form_values from existing object for unified template
    def _fmt(v):
        try:
            return (
                v.isoformat()
                if hasattr(v, "isoformat")
                else (v if v is not None else "")
            )
        except Exception:
            return v

    form_values = {
        "equipment_id": getattr(equipment_list, "equipment_id", ""),
        "equipment_code": getattr(equipment_list, "equipment_code", ""),
        "equipment_name_EN": getattr(equipment_list, "equipment_name_EN", ""),
        "equipment_name_TH": getattr(equipment_list, "equipment_name_TH", ""),
        "equipment_brand": getattr(equipment_list, "equipment_brand", ""),
        "equipment_model": getattr(equipment_list, "equipment_model", ""),
        "equipment_sn": getattr(equipment_list, "equipment_sn", ""),
        "equipment_gov": getattr(equipment_list, "equipment_gov", ""),
        "equipment_price": getattr(equipment_list, "equipment_price", ""),
        "equipment_photo": getattr(equipment_list, "equipment_photo", ""),
        "equipment_type": getattr(equipment_list, "equipment_type", ""),
        "equipment_life": getattr(equipment_list, "equipment_life", ""),
        "equipment_waranty_date": _fmt(
            getattr(equipment_list, "equipment_waranty_date", "")
        ),
        "equipment_waranty_due": _fmt(
            getattr(equipment_list, "equipment_waranty_due", "")
        ),
        "equipment_distributor_name": getattr(
            equipment_list, "equipment_distributor_name", ""
        ),
        "equipment_distributor_tel": getattr(
            equipment_list, "equipment_distributor_tel", ""
        ),
        "equipment_pm_fq": getattr(equipment_list, "equipment_pm_fq", ""),
        "equipment_pm_due": _fmt(getattr(equipment_list, "equipment_pm_due", "")),
        "equipment_cal_fq": getattr(equipment_list, "equipment_cal_fq", ""),
        "equipment_cal_due": _fmt(getattr(equipment_list, "equipment_cal_due", "")),
        "equipment_owner_customer": getattr(
            equipment_list, "equipment_owner_customer", ""
        ),
        "equipment_user_customer": getattr(
            equipment_list, "equipment_user_customer", ""
        ),
        "equipment_service_provider": getattr(
            equipment_list, "equipment_service_provider", ""
        ),
        "equipment_register_username": getattr(
            equipment_list, "equipment_register_username", ""
        ),
        "equipment_register_adminname": getattr(
            equipment_list, "equipment_register_adminname", ""
        ),
        "equipment_note": getattr(equipment_list, "equipment_note", ""),
    }

    # provide master lists for selects (with ancestors)
    def _ml(cat):
        qs = list(
            MasterItem.objects.filter(category=cat, active=True).order_by(
                "order", "label"
            )
        )
        if not qs:
            return []
        label_map = {m.id: m.label for m in qs}
        parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

        def build_ancestors(mid):
            chain = []
            seen = set()
            cur = parent_map.get(mid)
            while cur and cur not in seen:
                seen.add(cur)
                lbl = label_map.get(cur)
                if lbl:
                    chain.append(lbl)
                cur = parent_map.get(cur)
            chain.reverse()
            return " > ".join(chain)

        out = []
        for m in qs:
            out.append(
                {"id": m.id, "code": m.code, "label": m.label, "ancestors": build_ancestors(m.id)}
            )
        return out

    # Build master_equipments list with code, label, and meta fields
    def _ml_equipments():
        # Build lookup dictionaries for value -> label mapping
        # Use both code and label as keys (some master data use label as value)
        def build_lookup(items):
            d = {}
            for item in items:
                if item.get('code'):
                    d[item['code']] = item['label']
                d[item['label']] = item['label']  # label -> label mapping
            return d
        
        lookup_risk = build_lookup(_ml("risks"))
        lookup_type = build_lookup(_ml("equipment_type"))
        lookup_pm = build_lookup(_ml("frequency_pm"))
        lookup_cal = build_lookup(_ml("frequency_cal"))
        lookup_sv_provider = build_lookup(_ml("service_providers"))
        
        qs = list(
            MasterItem.objects.filter(category='equipments', active=True).order_by(
                "order", "code"
            )
        )
        
        out = []
        for m in qs:
            meta = m.meta or {}
            # Accept several possible meta key names for PM/CAL frequency
            pm_val = (
                meta.get("equipment_pm_fq")
                or meta.get("pm_fq")
                or meta.get("frequency_pm")
                or meta.get("frequency")
                or ""
            )
            cal_val = (
                meta.get("equipment_cal_fq")
                or meta.get("cal_fq")
                or meta.get("frequency_cal")
                or meta.get("frequency")
                or ""
            )
            risk_val = meta.get("equipment_risk", "")
            type_val = meta.get("equipment_type", "")
            lifespan_val = meta.get("equipment_life", "") or meta.get("lifespan", "")
            sv_provider_val = meta.get("equipment_service_provider", "") or meta.get("sv_provider", "")
            
            # Convert values to labels using lookup dictionaries
            risk_label = lookup_risk.get(risk_val, "")
            pm_fq_label = lookup_pm.get(pm_val, "")
            cal_fq_label = lookup_cal.get(cal_val, "")
            type_label = lookup_type.get(type_val, "")
            sv_provider_label = lookup_sv_provider.get(sv_provider_val, "")
            
            out.append({
                "id": m.id,
                "code": m.code,
                "label": m.label,
                "name_en": meta.get("equipment_name_en", ""),
                "name_th": meta.get("equipment_name_th", ""),
                # Additional master metadata used in equipment form/modal
                "risk": risk_val,
                "risk_label": risk_label,
                "pm_fq": pm_val,
                "pm_fq_label": pm_fq_label,
                "cal_fq": cal_val,
                "cal_fq_label": cal_fq_label,
                "equipment_type": type_val,
                "equipment_type_label": type_label,
                "lifespan": lifespan_val,
                "sv_provider": sv_provider_val,
                "sv_provider_label": sv_provider_label,
            })
        return out

    context = {
        "form_values": form_values,
        "edit": True,
        "equipment_list": equipment_list,
        "has_equipment_image": _equipment_image_exists(equipment_list),
        "master_equipment_type": _ml("equipment_type"),
        "master_frequency_cal": _ml("frequency_cal"),
        "master_frequency_pm": _ml("frequency_pm"),
        "master_manufacturers": _ml("manufacturers"),
        "master_departments": _ml("departments"),
        "master_units": _ml("units"),
        "master_customers": _ml("customers"),
        "master_locations": _ml("locations"),
        "master_service_providers": _ml("service_providers"),
        "master_equipments": _ml_equipments(),
    }
    return render(request, "equipment/equipment_form.html", context)


@permission_required("cmms.change_equipment_list", raise_exception=True)
def edit_equipment_by_code(request, code):
    """Redirect helper: accept equipment_id (code) and redirect to PK-based edit view.

    Keeps a single canonical edit implementation (edit_equipment) while allowing
    external links to reference equipment by code.
    """
    eq = get_object_or_404(Equipment_list, equipment_id=code)
    return redirect("edit_equipment", equipment_list_id=eq.id)


@login_required
def equipment_profile(request, code):
    """Equipment registration card view (ทะเบียนเครื่องมือ)."""
    from .models import WorkOrder

    equipment = get_object_or_404(Equipment_list, equipment_id=code)
    today = timezone.localdate()

    def build_due_state(due_date, required=True, warning_days=30):
        # ถ้าไม่มีวันครบกำหนด
        if not due_date:
            # ถ้าไม่ต้องทำ (เช่น ไม่ต้อง CAL)
            if not required:
                return {
                    "tone": "muted",
                    "text": "ไม่ต้องดำเนินการ",
                    "days": None,
                }
            # ถ้าต้องทำแต่ยังไม่ได้กำหนดวัน
            return {
                "tone": "muted",
                "text": "ยังไม่ระบุวันครบกำหนด",
                "days": None,
            }

        # มีวันครบกำหนด -> คำนวณสถานะ
        diff = (due_date - today).days
        if diff < 0:
            return {
                "tone": "danger",
                "text": f"เกินกำหนด {abs(diff)} วัน",
                "days": diff,
            }
        if diff <= warning_days:
            return {
                "tone": "warn",
                "text": f"ใกล้ครบกำหนด ({diff} วัน)",
                "days": diff,
            }
        return {
            "tone": "ok",
            "text": f"ปกติ (เหลือ {diff} วัน)",
            "days": diff,
        }

    workorders_qs = (
        WorkOrder.objects.filter(equipment=equipment)
        .select_related("assigned_to", "reported_by")
        .order_by("-reported_at")
    )

    total_workorders = workorders_qs.count()
    active_workorders = workorders_qs.exclude(
        status__in=["completed", "verified", "closed"]
    ).count()
    latest_workorders = list(workorders_qs[:8])
    qr_payload = request.build_absolute_uri(
        reverse("equipment_profile", kwargs={"code": equipment.equipment_id})
    )
    qr_image_url = (
        "https://api.qrserver.com/v1/create-qr-code/?size=220x220&margin=0&data="
        f"{quote(qr_payload)}"
    )

    context = {
        "equipment": equipment,
        "has_equipment_image": _equipment_image_exists(equipment),
        "today": today,
        "qr_payload": qr_payload,
        "qr_image_url": qr_image_url,
        "due_pm": build_due_state(
            equipment.equipment_pm_due,
            required=bool(equipment.requires_pm),
            warning_days=7,
        ),
        "due_cal": build_due_state(
            equipment.equipment_cal_due,
            required=bool(equipment.requires_cal),
            warning_days=14,
        ),
        "due_warranty": build_due_state(
            equipment.equipment_waranty_due,
            required=True,
            warning_days=30,
        ),
        "total_workorders": total_workorders,
        "active_workorders": active_workorders,
        "latest_workorders": latest_workorders,
    }
    return render(request, "equipment/equipment_profile.html", context)


@login_required
def equipment_profile_by_code(request, code):
    """Legacy route helper: redirect old by-code path to canonical profile path."""
    eq = get_object_or_404(Equipment_list, equipment_id=code)
    return redirect("equipment_profile", code=eq.equipment_id)


@login_required
def equipment_history(request, equipment_list_id):
    """Display registration, repair, and spare-parts history for equipment"""
    from .models import EquipmentHistory, WorkOrder, WorkOrderSparePart
    from .history_utils import get_equipment_history, format_field_name

    equipment = get_object_or_404(Equipment_list, id=equipment_list_id)
    history = get_equipment_history(equipment.equipment_id)

    # Format field names in history for display
    for record in history:
        if record.changed_fields:
            record.changed_fields = [format_field_name(f) for f in record.changed_fields]

    # Work orders for this equipment
    work_orders = (
        WorkOrder.objects.filter(equipment=equipment)
        .order_by("-reported_at")
        .prefetch_related("used_spare_parts__spare_part", "logs")
        .select_related("assigned_to", "reported_by", "verified_by")
    )

    # All spare-part entries used in WOs for this equipment
    spare_parts_qs = (
        WorkOrderSparePart.objects.filter(workorder__equipment=equipment)
        .select_related("spare_part", "workorder")
        .order_by("-workorder__reported_at", "id")
    )

    # Aggregate total cost
    total_sp_cost = sum(
        (float(sp.custom_unit_cost or 0) if sp.spare_part is None else float(sp.spare_part.unit_cost or 0)) * sp.quantity
        for sp in spare_parts_qs
    )

    wo_status_counts = {}
    for wo in work_orders:
        wo_status_counts[wo.status] = wo_status_counts.get(wo.status, 0) + 1

    context = {
        "equipment": equipment,
        "history": history,
        "work_orders": work_orders,
        "spare_parts_qs": spare_parts_qs,
        "wo_count": work_orders.count(),
        "sp_count": spare_parts_qs.count(),
        "total_sp_cost": total_sp_cost,
        "wo_status_counts": wo_status_counts,
    }
    return render(request, "equipment/equipment_history.html", context)


@permission_required("cmms.delete_equipment_list", raise_exception=True)
def delete_equipment_by_code(request, code):
    """Redirect helper: accept equipment_id (code) and redirect to PK-based delete view.

    Note: the target `delete_equipment` performs the delete action; this redirect
    preserves current behavior and permissions.
    """
    eq = get_object_or_404(Equipment_list, equipment_id=code)
    return redirect("delete_equipment", equipment_list_id=eq.id)


@permission_required("cmms.delete_equipment_list", raise_exception=True)
def delete_equipment(request, equipment_list_id):
    equipment_list = Equipment_list.objects.get(id=equipment_list_id)
    equipment_list.delete()
    messages.success(request, "ลบข้อมูลเรียบร้อยแล้ว")
    return redirect("/equipment_list")


# อัพโหลดไฟล์ Excel
@permission_required("cmms.add_equipment_list", raise_exception=True)
def equipment_download_template(request):
    """Generate and return an Excel .xlsx template file for Equipment import."""
    from django.http import HttpResponse
    from openpyxl.workbook import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = Workbook()
    ws = wb.active
    ws.title = "Equipment Template"

    # Headers based on Equipment_list model fields
    headers = [
        "equipment_id",
        "equipment_code",
        "equipment_name_TH",
        "equipment_name_EN",
        "equipment_brand",
        "equipment_model",
        "equipment_sn",
        "equipment_gov",
        "equipment_price",
        "equipment_photo",
        "equipment_type",
        "equipment_life",
        "equipment_waranty_date",
        "equipment_waranty_due",
        "equipment_distributor_name",
        "equipment_distributor_tel",
        "equipment_pm_fq",
        "equipment_pm_due",
        "equipment_cal_fq",
        "equipment_cal_due",
        "equipment_owner_customer",
        "equipment_user_customer",
        "equipment_register_username",
        "equipment_register_adminname",
        "equipment_note",
    ]

    # Style header row
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center")

    # Add headers
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        ws.column_dimensions[cell.column_letter].width = 20

    # Add sample row with example data
    sample_data = [
        "EQ-2024-001",  # equipment_id
        "EML-001",  # equipment_code (รหัส EML)
        "เครื่องวัดความดัน",  # equipment_name_TH
        "Blood Pressure Monitor",  # equipment_name_EN
        "Omron",  # equipment_brand
        "HEM-7120",  # equipment_model
        "SN123456789",  # equipment_sn
        "GOV-2024-001",  # equipment_gov
        "15000.00",  # equipment_price
        "",  # equipment_photo
        "Medical",  # equipment_type
        "10",  # equipment_life (years)
        "2024-01-01",  # equipment_waranty_date
        "2025-01-01",  # equipment_waranty_due
        "บริษัท ABC จำกัด",  # equipment_distributor_name
        "02-123-4567",  # equipment_distributor_tel
        "12",  # equipment_pm_fq (months)
        "2024-12-31",  # equipment_pm_due
        "12",  # equipment_cal_fq (months)
        "2024-12-31",  # equipment_cal_due
        "แผนกเวชศาสตร์ฉุกเฉิน",  # equipment_owner_customer
        "แผนกเวชศาสตร์ฉุกเฉิน",  # equipment_user_customer
        "",  # equipment_register_username
        "",  # equipment_register_adminname
        "ตัวอย่างข้อมูล",  # equipment_note
    ]

    ws.append(sample_data)

    # Set response
    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = 'attachment; filename="equipment_import_template.xlsx"'
    wb.save(response)
    return response


@require_http_methods(["POST"])
@permission_required("cmms.change_equipment_list", raise_exception=True)
def bulk_edit_equipment(request):
    """Bulk edit multiple equipment records"""
    import json
    from django.http import JsonResponse
    
    try:
        equipment_ids = json.loads(request.POST.get('equipment_ids', '[]'))
        
        if not equipment_ids:
            return JsonResponse({'success': False, 'error': 'ไม่มีรายการที่เลือก'})
        
        # Get fields to update
        updates = {}
        
        # EML Code and Equipment Names
        if request.POST.get('edit_code'):
            updates['equipment_code'] = request.POST.get('bulk_code', '').strip()
        
        if request.POST.get('edit_name_en'):
            updates['equipment_name_EN'] = request.POST.get('bulk_name_en', '').strip()
        
        if request.POST.get('edit_name_th'):
            updates['equipment_name_TH'] = request.POST.get('bulk_name_th', '').strip()
        
        # Serial Number and Gov Number
        if request.POST.get('edit_sn'):
            updates['equipment_sn'] = request.POST.get('bulk_sn', '').strip()
        
        if request.POST.get('edit_gov'):
            updates['equipment_gov'] = request.POST.get('bulk_gov', '').strip()
        
        # Basic info
        if request.POST.get('edit_type'):
            updates['equipment_type'] = request.POST.get('bulk_type', '').strip()
        
        if request.POST.get('edit_brand'):
            updates['equipment_brand'] = request.POST.get('bulk_brand', '').strip()
        
        if request.POST.get('edit_model'):
            updates['equipment_model'] = request.POST.get('bulk_model', '').strip()
        
        if request.POST.get('edit_life'):
            life = request.POST.get('bulk_life', '').strip()
            try:
                # Try to convert to int, stripping any non-numeric characters
                life_int = int(''.join(filter(str.isdigit, life))) if life else None
                updates['equipment_life'] = life_int
            except (ValueError, TypeError):
                updates['equipment_life'] = None
        
        if request.POST.get('edit_price'):
            price = request.POST.get('bulk_price', '').strip()
            updates['equipment_price'] = float(price) if price else None
        
        # Distributor info
        if request.POST.get('edit_distributor_name'):
            updates['equipment_distributor_name'] = request.POST.get('bulk_distributor_name', '').strip()
        
        if request.POST.get('edit_distributor_tel'):
            updates['equipment_distributor_tel'] = request.POST.get('bulk_distributor_tel', '').strip()
        
        # Owner/User info
        if request.POST.get('edit_owner'):
            updates['equipment_owner_customer'] = request.POST.get('bulk_owner', '').strip()
        
        if request.POST.get('edit_user'):
            updates['equipment_user_customer'] = request.POST.get('bulk_user', '').strip()
        
        if request.POST.get('edit_service_provider'):
            updates['equipment_service_provider'] = request.POST.get('bulk_service_provider', '').strip()
        
        # PM info
        if request.POST.get('edit_pm_fq'):
            pm_fq = request.POST.get('bulk_pm_fq', '').strip()
            try:
                pm_fq_int = int(''.join(filter(str.isdigit, pm_fq))) if pm_fq else None
                updates['equipment_pm_fq'] = pm_fq_int
            except (ValueError, TypeError):
                updates['equipment_pm_fq'] = None
        
        if request.POST.get('edit_pm_due'):
            pm_due = request.POST.get('bulk_pm_due', '').strip()
            updates['equipment_pm_due'] = pm_due if pm_due else None
        
        # CAL info
        if request.POST.get('edit_cal_fq'):
            cal_fq = request.POST.get('bulk_cal_fq', '').strip()
            try:
                cal_fq_int = int(''.join(filter(str.isdigit, cal_fq))) if cal_fq else None
                updates['equipment_cal_fq'] = cal_fq_int
            except (ValueError, TypeError):
                updates['equipment_cal_fq'] = None
        
        if request.POST.get('edit_cal_due'):
            cal_due = request.POST.get('bulk_cal_due', '').strip()
            updates['equipment_cal_due'] = cal_due if cal_due else None
        
        # Warranty info
        if request.POST.get('edit_warranty_date'):
            warranty_date = request.POST.get('bulk_warranty_date', '').strip()
            updates['equipment_waranty_date'] = warranty_date if warranty_date else None
        
        if request.POST.get('edit_warranty_due'):
            warranty_due = request.POST.get('bulk_warranty_due', '').strip()
            updates['equipment_waranty_due'] = warranty_due if warranty_due else None
        
        # Note
        if request.POST.get('edit_note'):
            updates['equipment_note'] = request.POST.get('bulk_note', '').strip()
        
        if not updates:
            return JsonResponse({'success': False, 'error': 'ไม่มีฟิลด์ที่เลือกแก้ไข'})
        
        # Update equipment records
        updated_count = Equipment_list.objects.filter(id__in=equipment_ids).update(**updates)
        
        return JsonResponse({
            'success': True,
            'updated_count': updated_count,
            'message': f'อัพเดทข้อมูลสำเร็จ {updated_count} รายการ'
        })
        
    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)})


def import_equipment_list(request):
    if pd is None:
        messages.error(
            request, "ติดต่อ admin: ไลบรารี pandas ยังไม่ได้ติดตั้ง (required for Excel import)"
        )
        return redirect("equipment_list")

    if request.method == "POST":
        excel_file = request.FILES["excel_file"]
        df = pd.read_excel(excel_file)
        
        # Helper function to convert date values
        def convert_date(value):
            """Convert various date formats to string or None"""
            if pd.isna(value):
                return None
            if isinstance(value, pd.Timestamp):
                return value.strftime('%Y-%m-%d')
            if isinstance(value, str):
                return value if value.strip() else None
            return None
        
        # Helper function to convert numeric values
        def convert_number(value):
            """Convert numeric values, return None for NaN"""
            if pd.isna(value):
                return None
            return value

        for _, row in df.iterrows():
            # Only pass fields that exist on the current Equipment_list model.
            # Legacy columns (sv_*, owner_/user_* detail fields, pm_date/cal_date)
            # were removed in the model; passing them as kwargs causes errors.
            Equipment_list.objects.create(
                equipment_id=row.get("equipment_id"),
                equipment_code=row.get("equipment_code"),
                equipment_name_TH=row.get("equipment_name_TH"),
                equipment_name_EN=row.get("equipment_name_EN"),
                equipment_brand=row.get("equipment_brand"),
                equipment_model=row.get("equipment_model"),
                equipment_sn=row.get("equipment_sn"),
                equipment_gov=row.get("equipment_gov"),
                equipment_price=convert_number(row.get("equipment_price")),
                equipment_photo=row.get("equipment_photo"),
                equipment_type=row.get("equipment_type"),
                equipment_life=convert_number(row.get("equipment_life")),
                equipment_waranty_date=convert_date(row.get("equipment_waranty_date")),
                equipment_waranty_due=convert_date(row.get("equipment_waranty_due")),
                equipment_distributor_name=row.get("equipment_distributor_name"),
                equipment_distributor_tel=row.get("equipment_distributor_tel"),
                equipment_pm_fq=convert_number(row.get("equipment_pm_fq")),
                equipment_pm_due=convert_date(row.get("equipment_pm_due")),
                equipment_cal_fq=convert_number(row.get("equipment_cal_fq")),
                equipment_cal_due=convert_date(row.get("equipment_cal_due")),
                equipment_owner_customer=row.get("equipment_owner_customer"),
                equipment_user_customer=row.get("equipment_user_customer"),
                equipment_register_username=row.get("equipment_register_username"),
                equipment_register_adminname=row.get("equipment_register_adminname"),
                equipment_note=row.get("equipment_note"),
            )
        messages.success(request, "นำเข้าข้อมูลเรียบร้อยแล้ว")
        return redirect("equipment_list")
    return render(request, "equipment/import_equipment_list.html")


# ดาวน์โหลดไฟล์ Excel
@permission_required("cmms.view_equipment_list", raise_exception=True)
def export_equipment_list(request):
    q = request.GET.get("q", "").strip()
    equipment_name_ENs_selected = request.GET.getlist("equipment_name_EN")
    equipment_name_THs_selected = request.GET.getlist("equipment_name_TH")
    equipment_brands_selected = request.GET.getlist("equipment_brand")
    equipment_models_selected = request.GET.getlist("equipment_model")
    # Legacy fields removed - use equipment_user_customer only
    equipment_user_customers_selected = request.GET.getlist("equipment_user_customer")
    all_equipment = Equipment_list.objects.all()
    if q:
        from django.db.models import Q

        all_equipment = all_equipment.filter(
            Q(equipment_id__icontains=q)
            | Q(equipment_name_EN__icontains=q)
            | Q(equipment_name_TH__icontains=q)
            | Q(equipment_brand__icontains=q)
            | Q(equipment_model__icontains=q)
            | Q(equipment_sn__icontains=q)
            | Q(equipment_user_customer__icontains=q)
            | Q(equipment_owner_customer__icontains=q)
        )
    if equipment_name_ENs_selected and any(
        val for val in equipment_name_ENs_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_name_EN__in=[val for val in equipment_name_ENs_selected if val]
        )
    if equipment_name_THs_selected and any(
        val for val in equipment_name_THs_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_name_TH__in=[val for val in equipment_name_THs_selected if val]
        )
    if equipment_brands_selected and any(
        val for val in equipment_brands_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_brand__in=[val for val in equipment_brands_selected if val]
        )
    if equipment_models_selected and any(
        val for val in equipment_models_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_model__in=[val for val in equipment_models_selected if val]
        )
    # Filter by equipment_user_customer only (legacy fields removed)
    if equipment_user_customers_selected and any(
        val for val in equipment_user_customers_selected if val
    ):
        all_equipment = all_equipment.filter(
            equipment_user_customer__in=[
                val for val in equipment_user_customers_selected if val
            ]
        )

    if pd is None:
        return HttpResponse(
            "Server error: pandas not installed (required to export Excel).", status=500
        )

    equipment_list = all_equipment.values()
    df = pd.DataFrame(equipment_list)

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    response["Content-Disposition"] = "attachment; filename=equipment_list.xlsx"

    with pd.ExcelWriter(response, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Equipment_list")

    return response


# WorkOrder views
@login_required
@permission_required("cmms.view_workorder", raise_exception=True)
def workorder_list(request):
    from django.db.models import Q

    from cmms.models import MasterItem, WorkOrder

    q = request.GET.get("q", "").strip()
    all_wos = WorkOrder.objects.select_related(
        "equipment", "assigned_to", "origin_request"
    ).prefetch_related("workorder_types")
    if q:
        all_wos = all_wos.filter(title__icontains=q)

    # filter by status (comma-separated) or special keyword 'pending'
    status_param = request.GET.get("status", "").strip()
    if status_param:
        if status_param == "pending":
            all_wos = all_wos.filter(status__in=["open", "assigned", "in_progress"])
        else:
            statuses = [s.strip() for s in status_param.split(",") if s.strip()]
            if statuses:
                all_wos = all_wos.filter(status__in=statuses)

    work_type_param = request.GET.get("work_type", "").strip()
    work_type_options = []
    work_type_items = list(
        MasterItem.objects.filter(category="workorder_type", active=True).order_by(
            "order", "label"
        )
    )
    for item in work_type_items:
        work_type_options.append(
            {"value": f"master:{item.pk}", "label": item.label}
        )

    master_type_codes = {
        item.code.strip().casefold()
        for item in work_type_items
        if item.code and item.code.strip()
    }
    for value, label in WorkOrder.WORKORDER_TYPE_CHOICES:
        if value.casefold() not in master_type_codes:
            work_type_options.append(
                {"value": f"legacy:{value}", "label": label}
            )

    if work_type_param.startswith("master:"):
        try:
            selected_master_type = next(
                item
                for item in work_type_items
                if str(item.pk) == work_type_param.removeprefix("master:")
            )
        except StopIteration:
            selected_master_type = None
        if selected_master_type:
            legacy_values = [selected_master_type.label]
            if selected_master_type.code:
                legacy_values.append(selected_master_type.code)
            type_query = Q(workorder_types=selected_master_type) | Q(
                workorder_type__in=legacy_values
            )
            if selected_master_type.label:
                type_query |= Q(title__icontains=selected_master_type.label)
            if selected_master_type.code:
                type_query |= Q(title__icontains=selected_master_type.code)
            all_wos = all_wos.filter(type_query)
        else:
            work_type_param = ""
    elif work_type_param.startswith("legacy:"):
        legacy_type = work_type_param.removeprefix("legacy:")
        if legacy_type in dict(WorkOrder.WORKORDER_TYPE_CHOICES):
            legacy_values = [legacy_type]
            selected_master_types = [
                item
                for item in work_type_items
                if item.code.casefold() == legacy_type.casefold()
                or item.label.casefold() == legacy_type.casefold()
            ]
            type_query = Q(workorder_type__iexact=legacy_type) | Q(
                workorder_types__in=selected_master_types
            )
            type_labels = {
                item.label for item in selected_master_types if item.label
            }
            type_labels.update(
                label
                for value, label in WorkOrder.WORKORDER_TYPE_CHOICES
                if value.casefold() == legacy_type.casefold()
            )
            for label in type_labels:
                type_query |= Q(title__icontains=label)
            type_query |= Q(title__icontains=legacy_type)
            all_wos = all_wos.filter(type_query)
        else:
            work_type_param = ""
    else:
        work_type_param = ""

    # filter by today
    today_param = request.GET.get("today", "").lower()
    if today_param in ("1", "true", "yes"):
        from django.utils import timezone

        today_date = timezone.localtime().date()
        all_wos = all_wos.filter(reported_at__date=today_date)
    return render(
        request,
        "workorder/workorder_list.html",
        {
            "workorders": all_wos.distinct(),
            "q": q,
            "work_type_options": work_type_options,
            "selected_work_type": work_type_param,
        },
    )


@login_required
@permission_required("cmms.add_workorder", raise_exception=True)
def workorder_create(request):
    from cmms.models import ServiceRequest, ServiceRequestLog

    from .forms import WorkOrderForm

    if request.method == "POST":
        form = WorkOrderForm(request.POST)
        # Enforce: workorder must originate from a Service Request
        origin_sr_val = request.POST.get("origin_sr")
        if not origin_sr_val:
            messages.error(request, "ต้องสร้างใบงานจากคำขอรับบริการเท่านั้น")
            return redirect("service_request_create")
        # Re-lookup the SR so context is preserved if form validation fails
        sr_obj = None
        try:
            sr_obj = ServiceRequest.objects.get(id=int(origin_sr_val))
        except Exception:
            pass
        if form.is_valid():
            wo = form.save(commit=False)
            wo.reported_by = request.user
            # capture whether the user accepted the suggested title (hidden input)
            try:
                accepted = request.POST.get("accepted_title_suggestion", "0")
                wo.accepted_title_suggestion = (
                    True if str(accepted) in ("1", "true", "True") else False
                )
            except Exception:
                wo.accepted_title_suggestion = False
            wo.save()
            form.save_m2m()
            # If this workorder was created from a ServiceRequest (pre-fill flow), link them and add logs
            origin_sr = request.POST.get("origin_sr")
            if origin_sr:
                try:
                    sr = ServiceRequest.objects.get(id=int(origin_sr))
                    sr.converted_to = wo
                    sr.status = "converted"
                    sr.save()
                    ServiceRequestLog.objects.create(
                        request=sr,
                        action="converted_via_create_form",
                        actor=request.user,
                        note=f"Converted to WO#{wo.id} via create form",
                    )
                    from cmms.models import WorkOrderLog

                    WorkOrderLog.objects.create(
                        workorder=wo,
                        action="created_from_request",
                        actor=request.user,
                        note=f"From SR#{sr.id}",
                    )
                except Exception:
                    pass
            # After creating a workorder, redirect to workorder list
            messages.success(request, f"สร้างใบงาน WO#{wo.id} สำเร็จ")
            return redirect("workorder_list")
    else:
        # support prefill from ServiceRequest: ?from_sr=<id>
        form = None
        sr_obj = None
        from_sr = request.GET.get("from_sr")
        
        # Support prefill from equipment: ?equipment_id=<qr_code>
        equipment_id_param = request.GET.get("equipment_id")
        equipment_obj = None
        
        if from_sr:
            try:
                sr_obj = ServiceRequest.objects.get(id=int(from_sr))
                initial = {
                    "title": sr_obj.title,
                    "description": sr_obj.description,
                    "equipment": sr_obj.equipment.id if sr_obj.equipment else None,
                }
                form = WorkOrderForm(initial=initial)
            except Exception:
                form = WorkOrderForm()
        else:
            # Work orders must be created from a Service Request
            messages.info(request, "การสร้างใบงานต้องเริ่มจากคำขอรับบริการ กรุณาสร้างคำขอก่อน")
            return redirect("service_request_create")

    # Provide active technicians list for workorder create page (for easier assignment)
    try:
        from .models import Technician
        technicians_qs = Technician.objects.filter(active=True).order_by("name")
    except Exception:
        technicians_qs = []

    ctx = {"form": form, "technicians": technicians_qs}
    if "sr_obj" in locals() and sr_obj:
        ctx["sr"] = sr_obj
        ctx["origin_sr"] = sr_obj.id
    if "equipment_obj" in locals() and equipment_obj:
        ctx["prefill_equipment"] = equipment_obj
    return render(request, "workorder/workorder_create.html", ctx)


@permission_required("cmms.change_workorder", raise_exception=True)
def workorder_detail(request, wo_id):
    from cmms.models import WorkOrder, WorkOrderAttachment, WorkOrderComment

    wo = WorkOrder.objects.get(id=wo_id)
    if request.method == "POST":
        # handle workflow action buttons
        action = request.POST.get("action")
        if action == "accept_request" and request.user.has_perm(
            "cmms.change_workorder"
        ):
            wo.accepted_by = request.user
            from django.utils import timezone

            wo.accepted_at = timezone.now()
            wo.accepted_note = request.POST.get("accepted_note", "")
            wo.status = "assigned"
            wo.save()
            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo,
                action="accepted",
                actor=request.user,
                note=wo.accepted_note,
            )
            messages.success(request, "รับคำขอแล้ว และรอตั้งค่าเพื่อมอบหมายงาน")
            return redirect("workorder_detail", wo_id=wo.id)

        if action == "assign_to_tech" and request.user.has_perm(
            "cmms.change_workorder"
        ):
            assigned_to_id = request.POST.get("assigned_to")
            try:
                from django.contrib.auth import get_user_model

                User = get_user_model()
                if assigned_to_id:
                    wo.assigned_to = User.objects.get(id=int(assigned_to_id))
                    wo.status = "assigned"
                    wo.save()
                    from cmms.models import WorkOrderLog

                    WorkOrderLog.objects.create(
                        workorder=wo,
                        action="assigned",
                        actor=request.user,
                        note=f"Assigned to {wo.assigned_to}",
                    )
                    messages.success(request, "มอบหมายงานเรียบร้อยแล้ว")
            except Exception:
                messages.error(request, "ไม่สามารถมอบหมายงานได้")
            return redirect("workorder_detail", wo_id=wo.id)

        # technician actions
        if action == "tech_start" and request.user == wo.assigned_to:
            from django.utils import timezone

            wo.actual_start = timezone.now()
            wo.status = "in_progress"
            wo.save()
            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo, action="started", actor=request.user
            )
            messages.success(request, "ช่างเริ่มงานแล้ว")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-tech")

        if action == "tech_complete" and request.user == wo.assigned_to:
            from django.utils import timezone

            from .models import MasterItem

            wo.actual_end = timezone.now()
            wo.status = "completed"

            # Collect separated technician inputs and combine them into technician_report
            cause = request.POST.get("technician_cause", "").strip()
            analysis = request.POST.get("technician_analysis", "").strip()
            method_id = request.POST.get("technician_method", "").strip()
            actions_done = request.POST.get("technician_actions", "").strip()

            method_label = ""
            if method_id:
                try:
                    method_item = MasterItem.objects.get(id=int(method_id))
                    method_label = method_item.label
                except Exception:
                    method_label = ""

            parts = []
            if cause:
                parts.append(f"สาเหตุ: {cause}")
            if analysis:
                parts.append(f"วิเคราะห์สาเหตุ: {analysis}")
            if method_label:
                parts.append(f"วิธีการดำเนินงาน: {method_label}")
            if actions_done:
                parts.append(f"การดำเนินงาน: {actions_done}")

            wo.technician_report = "\n\n".join(parts)
            wo.save()

            # Save spare parts used in this work order
            from .models import SparePart, WorkOrderSparePart
            import json as _json

            spare_parts_json = request.POST.get("spare_parts_data", "[]")
            try:
                spare_items = _json.loads(spare_parts_json)
            except (ValueError, TypeError):
                spare_items = []

            # Clear previous entries and re-create
            wo.used_spare_parts.all().delete()
            for item in spare_items:
                sp_id = item.get("id")
                qty = item.get("qty", 1)
                if not isinstance(qty, int) or qty < 1:
                    continue

                # Custom spare part (อะไหล่อื่นๆ)
                if sp_id == "custom":
                    c_name = str(item.get("custom_name", "")).strip()
                    if not c_name:
                        continue
                    c_cost = item.get("custom_unit_cost")
                    c_unit = str(item.get("custom_unit", "ชิ้น")).strip() or "ชิ้น"
                    try:
                        c_cost = round(float(c_cost), 2) if c_cost else None
                    except (ValueError, TypeError):
                        c_cost = None
                    WorkOrderSparePart.objects.create(
                        workorder=wo,
                        spare_part=None,
                        custom_name=c_name,
                        custom_unit_cost=c_cost,
                        custom_unit=c_unit,
                        quantity=qty,
                        is_out_of_stock=True,
                    )
                    continue

                # Normal spare part from inventory
                if not sp_id:
                    continue
                try:
                    sp = SparePart.objects.get(id=int(sp_id), active=True)
                    out_of_stock = sp.quantity < qty
                    if sp.quantity >= qty:
                        sp.quantity -= qty
                        sp.save(update_fields=["quantity"])
                    WorkOrderSparePart.objects.create(
                        workorder=wo, spare_part=sp, quantity=qty,
                        is_out_of_stock=out_of_stock,
                    )
                except SparePart.DoesNotExist:
                    continue

            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo,
                action="completed",
                actor=request.user,
                note=wo.technician_report,
            )
            messages.success(request, "ช่างบันทึกผลการดำเนินงานเรียบร้อยแล้ว")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-detail")

        # technician: put on hold (รออะไหล่ / รอ Outsource)
        if action == "tech_hold" and request.user == wo.assigned_to and wo.status == "in_progress":
            from django.utils import timezone
            from .models import WorkOrderLog
            import json as _json

            hold_reason = request.POST.get("hold_reason", "waiting_parts")
            hold_note = request.POST.get("hold_note", "").strip()

            # Save spare parts first (same logic as tech_complete)
            from .models import SparePart, WorkOrderSparePart
            spare_parts_json = request.POST.get("spare_parts_data", "[]")
            try:
                spare_items = _json.loads(spare_parts_json)
            except (ValueError, TypeError):
                spare_items = []
            wo.used_spare_parts.all().delete()
            for item in spare_items:
                sp_id = item.get("id")
                qty = item.get("qty", 1)
                if not isinstance(qty, int) or qty < 1:
                    continue
                if sp_id == "custom":
                    c_name = str(item.get("custom_name", "")).strip()
                    if not c_name:
                        continue
                    c_cost = item.get("custom_unit_cost")
                    c_unit = str(item.get("custom_unit", "ชิ้น")).strip() or "ชิ้น"
                    try:
                        c_cost = round(float(c_cost), 2) if c_cost else None
                    except (ValueError, TypeError):
                        c_cost = None
                    WorkOrderSparePart.objects.create(
                        workorder=wo, spare_part=None, custom_name=c_name,
                        custom_unit_cost=c_cost, custom_unit=c_unit,
                        quantity=qty, is_out_of_stock=True,
                    )
                    continue
                if not sp_id:
                    continue
                try:
                    sp = SparePart.objects.get(id=int(sp_id), active=True)
                    out_of_stock = sp.quantity < qty
                    if sp.quantity >= qty:
                        sp.quantity -= qty
                        sp.save(update_fields=["quantity"])
                    WorkOrderSparePart.objects.create(
                        workorder=wo, spare_part=sp, quantity=qty,
                        is_out_of_stock=out_of_stock,
                    )
                except SparePart.DoesNotExist:
                    continue

            # Save technician report
            cause = request.POST.get("technician_cause", "").strip()
            analysis = request.POST.get("technician_analysis", "").strip()
            actions_done = request.POST.get("technician_actions", "").strip()
            parts = []
            if cause:
                parts.append(f"สาเหตุ: {cause}")
            if analysis:
                parts.append(f"วิเคราะห์สาเหตุ: {analysis}")
            if actions_done:
                parts.append(f"การดำเนินงาน: {actions_done}")
            if parts:
                wo.technician_report = "\n\n".join(parts)

            # Set on_hold
            wo.status = "on_hold"
            wo.hold_reason = hold_reason
            wo.hold_note = hold_note
            # Outsource details
            if hold_reason == "waiting_outsource":
                wo.outsource_vendor = request.POST.get("outsource_vendor", "").strip()
                try:
                    wo.outsource_cost = float(request.POST.get("outsource_cost", "") or 0) or None
                except (ValueError, TypeError):
                    wo.outsource_cost = None
                exp_date = request.POST.get("outsource_expected_date", "").strip()
                if exp_date:
                    try:
                        from datetime import date as _date
                        wo.outsource_expected_date = _date.fromisoformat(exp_date)
                    except ValueError:
                        pass
            wo.save()

            reason_labels = {"waiting_parts": "รออะไหล่", "waiting_outsource": "รอผู้ให้บริการภายนอก", "other": "อื่นๆ"}
            log_note = f"ระงับชั่วคราว — {reason_labels.get(hold_reason, hold_reason)}"
            if hold_note:
                log_note += f": {hold_note}"
            WorkOrderLog.objects.create(workorder=wo, action="on_hold", actor=request.user, note=log_note)
            messages.warning(request, f"ระงับใบงานชั่วคราว — {reason_labels.get(hold_reason, hold_reason)}")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-detail")

        # technician: complete but with follow-up needed (เครื่องยังใช้งานได้ แต่มีอะไหล่ค้าง)
        if action == "tech_complete_followup" and request.user == wo.assigned_to and wo.status == "in_progress":
            from django.utils import timezone
            from .models import WorkOrderLog
            import json as _json

            wo.actual_end = timezone.now()
            wo.status = "completed"
            wo.follow_up_needed = True

            # Save technician report
            from .models import MasterItem
            cause = request.POST.get("technician_cause", "").strip()
            analysis = request.POST.get("technician_analysis", "").strip()
            method_id = request.POST.get("technician_method", "").strip()
            actions_done = request.POST.get("technician_actions", "").strip()
            method_label = ""
            if method_id:
                try:
                    method_item = MasterItem.objects.get(id=int(method_id))
                    method_label = method_item.label
                except Exception:
                    method_label = ""
            parts = []
            if cause:
                parts.append(f"สาเหตุ: {cause}")
            if analysis:
                parts.append(f"วิเคราะห์สาเหตุ: {analysis}")
            if method_label:
                parts.append(f"วิธีการดำเนินงาน: {method_label}")
            if actions_done:
                parts.append(f"การดำเนินงาน: {actions_done}")
            wo.technician_report = "\n\n".join(parts)
            wo.save()

            # Save spare parts (same logic)
            from .models import SparePart, WorkOrderSparePart
            spare_parts_json = request.POST.get("spare_parts_data", "[]")
            try:
                spare_items = _json.loads(spare_parts_json)
            except (ValueError, TypeError):
                spare_items = []
            wo.used_spare_parts.all().delete()
            for item in spare_items:
                sp_id = item.get("id")
                qty = item.get("qty", 1)
                if not isinstance(qty, int) or qty < 1:
                    continue
                if sp_id == "custom":
                    c_name = str(item.get("custom_name", "")).strip()
                    if not c_name:
                        continue
                    c_cost = item.get("custom_unit_cost")
                    c_unit = str(item.get("custom_unit", "ชิ้น")).strip() or "ชิ้น"
                    try:
                        c_cost = round(float(c_cost), 2) if c_cost else None
                    except (ValueError, TypeError):
                        c_cost = None
                    WorkOrderSparePart.objects.create(
                        workorder=wo, spare_part=None, custom_name=c_name,
                        custom_unit_cost=c_cost, custom_unit=c_unit,
                        quantity=qty, is_out_of_stock=True,
                    )
                    continue
                if not sp_id:
                    continue
                try:
                    sp = SparePart.objects.get(id=int(sp_id), active=True)
                    out_of_stock = sp.quantity < qty
                    if sp.quantity >= qty:
                        sp.quantity -= qty
                        sp.save(update_fields=["quantity"])
                    WorkOrderSparePart.objects.create(
                        workorder=wo, spare_part=sp, quantity=qty,
                        is_out_of_stock=out_of_stock,
                    )
                except SparePart.DoesNotExist:
                    continue

            WorkOrderLog.objects.create(
                workorder=wo, action="completed_follow_up", actor=request.user,
                note="เสร็จสิ้น (เครื่องใช้งานได้) — รอติดตั้งอะไหล่ภายหลัง",
            )
            messages.success(request, "บันทึกผลงาน — เครื่องใช้งานได้ รอติดตั้งอะไหล่ภายหลัง")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-detail")

        # officer: resume from on_hold
        if action == "resume_from_hold" and request.user.has_perm("cmms.change_workorder") and wo.status == "on_hold":
            from .models import WorkOrderLog
            wo.status = "in_progress"
            resume_note = request.POST.get("resume_note", "").strip()
            wo.hold_reason = ""
            wo.hold_note = ""
            wo.save()
            WorkOrderLog.objects.create(
                workorder=wo, action="resumed", actor=request.user,
                note=resume_note or "คืนสถานะ — ดำเนินงานต่อ",
            )
            messages.success(request, "คืนสถานะใบงาน — ดำเนินงานต่อ")
            return redirect("workorder_detail", wo_id=wo.id)

        # officer: cancel hold and close / change plan
        if action == "cancel_hold" and request.user.has_perm("cmms.change_workorder") and wo.status == "on_hold":
            from .models import WorkOrderLog
            cancel_target = request.POST.get("cancel_target", "in_progress")
            cancel_note = request.POST.get("cancel_note", "").strip()
            if cancel_target == "closed":
                wo.status = "closed"
                wo.hold_reason = ""
                wo.hold_note = ""
                wo.save()
                WorkOrderLog.objects.create(
                    workorder=wo, action="hold_cancelled_closed", actor=request.user,
                    note=cancel_note or "ยกเลิกการรอ — ปิดใบงาน",
                )
                messages.info(request, "ยกเลิกการรอ — ปิดใบงาน")
            else:
                wo.status = "in_progress"
                wo.hold_reason = ""
                wo.hold_note = ""
                wo.save()
                WorkOrderLog.objects.create(
                    workorder=wo, action="hold_cancelled_resumed", actor=request.user,
                    note=cancel_note or "ยกเลิกการรอ — เปลี่ยนแผน ดำเนินงานต่อ",
                )
                messages.info(request, "ยกเลิกการรอ — เปลี่ยนแผน ดำเนินงานต่อ")
            return redirect("workorder_detail", wo_id=wo.id)

        # fulfill (receive) a previously out-of-stock spare part
        if action == "fulfill_spare_part" and request.user.has_perm(
            "cmms.change_workorder"
        ):
            from .models import SparePart, WorkOrderSparePart, WorkOrderLog

            usp_id = request.POST.get("usp_id")
            try:
                usp = WorkOrderSparePart.objects.select_related("spare_part").get(
                    id=int(usp_id), workorder=wo, is_out_of_stock=True,
                )
                if usp.spare_part:
                    # อะไหล่จากคลัง — หักสต็อก
                    sp = usp.spare_part
                    if sp.quantity >= usp.quantity:
                        sp.quantity -= usp.quantity
                        sp.save(update_fields=["quantity"])
                        usp.is_out_of_stock = False
                        usp.save(update_fields=["is_out_of_stock"])
                        WorkOrderLog.objects.create(
                            workorder=wo,
                            action="spare_part_fulfilled",
                            actor=request.user,
                            note=f"รับอะไหล่ {sp.code} ({sp.name}) x{usp.quantity} — หักสต็อกแล้ว",
                        )
                        messages.success(
                            request,
                            f"รับอะไหล่ {sp.code} เรียบร้อย — หักสต็อกแล้ว (คงเหลือ {sp.quantity})",
                        )
                    else:
                        messages.warning(
                            request,
                            f"สต็อก {sp.code} ยังไม่เพียงพอ (คงเหลือ {sp.quantity}, ต้องการ {usp.quantity}) — กรุณาเติมสต็อกก่อน",
                        )
                else:
                    # อะไหล่อื่นๆ — แค่เปลี่ยนสถานะ ไม่ต้องหักสต็อก
                    usp.is_out_of_stock = False
                    usp.save(update_fields=["is_out_of_stock"])
                    WorkOrderLog.objects.create(
                        workorder=wo,
                        action="spare_part_fulfilled",
                        actor=request.user,
                        note=f"รับอะไหล่อื่นๆ '{usp.custom_name}' x{usp.quantity} — จัดซื้อเรียบร้อย",
                    )
                    messages.success(
                        request,
                        f"รับอะไหล่ '{usp.custom_name}' เรียบร้อย",
                    )
            except (WorkOrderSparePart.DoesNotExist, ValueError, TypeError):
                messages.error(request, "ไม่พบรายการอะไหล่ที่ต้องการ")

            # Check if ALL spare parts are now fulfilled — auto actions
            remaining = wo.used_spare_parts.filter(is_out_of_stock=True).count()
            if remaining == 0:
                if wo.status == "on_hold" and wo.hold_reason == "waiting_parts":
                    # Auto-resume: all parts arrived → back to in_progress
                    wo.status = "in_progress"
                    wo.hold_reason = ""
                    wo.hold_note = ""
                    wo.save()
                    WorkOrderLog.objects.create(
                        workorder=wo, action="auto_resumed",
                        actor=request.user,
                        note="อะไหล่ครบ — คืนสถานะอัตโนมัติ",
                    )
                    messages.success(request, "อะไหล่ครบแล้ว — ใบงานกลับเป็น ดำเนินงาน อัตโนมัติ")
                elif wo.follow_up_needed:
                    # Auto-create follow-up WO
                    follow_wo = WorkOrder.objects.create(
                        title=f"ติดตั้งอะไหล่ — WO#{wo.id} {wo.title}",
                        workorder_type=wo.workorder_type,
                        description=f"ใบงานติดตาม: ติดตั้งอะไหล่ที่จัดซื้อสำเร็จจาก WO#{wo.id}\n\nอะไหล่ที่ต้องติดตั้ง:\n" +
                            "\n".join(f"- {u.display_name} x{u.quantity}" for u in wo.used_spare_parts.all()),
                        equipment=wo.equipment,
                        reported_by=request.user,
                        assigned_to=wo.assigned_to,
                        priority=wo.priority,
                        status="assigned" if wo.assigned_to else "open",
                        follow_up_for=wo,
                    )
                    wo.follow_up_needed = False
                    wo.save(update_fields=["follow_up_needed"])
                    WorkOrderLog.objects.create(
                        workorder=wo, action="follow_up_created",
                        actor=request.user,
                        note=f"อะไหล่ครบ — สร้างใบงานติดตาม WO#{follow_wo.id}",
                    )
                    WorkOrderLog.objects.create(
                        workorder=follow_wo, action="created",
                        actor=request.user,
                        note=f"สร้างจากใบงานต้นเรื่อง WO#{wo.id}",
                    )
                    messages.success(
                        request,
                        f"อะไหล่ครบแล้ว — สร้างใบงานติดตาม WO#{follow_wo.id} สำเร็จ",
                    )

            return redirect("workorder_detail", wo_id=wo.id)

        # engineer verification
        if action == "engineer_verify" and request.user.has_perm(
            "cmms.change_workorder"
        ):
            from django.utils import timezone

            wo.verified_by = request.user
            wo.verified_at = timezone.now()
            wo.verified_note = request.POST.get("verified_note", "")
            wo.status = "verified"
            wo.save()
            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo,
                action="verified",
                actor=request.user,
                note=wo.verified_note,
            )
            messages.success(request, "วิศวกรยืนยันผลงานเรียบร้อยแล้ว")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-detail")

        # engineer rejection — send back to in_progress
        if action == "engineer_reject" and request.user.has_perm(
            "cmms.change_workorder"
        ):
            reject_note = request.POST.get("verified_note", "").strip()
            wo.status = "in_progress"
            wo.actual_end = None
            wo.save()
            from cmms.models import WorkOrderLog

            WorkOrderLog.objects.create(
                workorder=wo,
                action="rejected",
                actor=request.user,
                note=reject_note or "ส่งกลับแก้ไข",
            )
            messages.warning(request, "ส่งกลับให้ช่างแก้ไขเรียบร้อยแล้ว")
            from django.urls import reverse
            return redirect(reverse("workorder_detail", args=[wo.id]) + "#tab-detail")

        # update assigned user and status
        new_status = request.POST.get("status", wo.status)
        assigned_to_id = request.POST.get("assigned_to")
        notes = request.POST.get("notes", "")
        # set assigned_to if provided
        if assigned_to_id:
            try:
                from django.contrib.auth import get_user_model

                User = get_user_model()
                wo.assigned_to = User.objects.get(id=int(assigned_to_id))
            except Exception:
                pass
        wo.status = new_status
        wo.notes = notes
        wo.save()

        # handle comment
        comment_text = request.POST.get("comment", "").strip()
        if comment_text:
            WorkOrderComment.objects.create(
                workorder=wo, author=request.user, text=comment_text
            )

        # handle file upload
        if request.FILES.get("attachment"):
            f = request.FILES["attachment"]
            attachment_name = request.POST.get("attachment_name", "").strip()
            attachment_type = request.POST.get("attachment_type", "other").strip() or "other"
            # Validate file type and size
            import os
            allowed_extensions = {".pdf", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".doc", ".docx", ".xls", ".xlsx"}
            max_upload_size = 10 * 1024 * 1024  # 10MB
            ext = os.path.splitext(f.name)[1].lower()
            allowed_document_types = {v for v, _ in WorkOrderAttachment.DOCUMENT_TYPE_CHOICES}
            if attachment_type not in allowed_document_types:
                attachment_type = "other"
            if ext not in allowed_extensions:
                messages.error(request, f"ไฟล์ประเภท {ext} ไม่รองรับ รองรับ: PDF, รูปภาพ, Word, Excel")
            elif f.size > max_upload_size:
                messages.error(request, f"ไฟล์ขนาดใหญ่เกินไป ({f.size // (1024*1024)}MB) สูงสุด 10MB")
            else:
                WorkOrderAttachment.objects.create(
                    workorder=wo,
                    uploaded_by=request.user,
                    file=f,
                    document_name=attachment_name,
                    document_type=attachment_type,
                )

        # notify via Celery task (if available)
        try:
            from .tasks import send_assignment_email

            if wo.assigned_to and wo.assigned_to.email:
                send_assignment_email.delay(
                    wo.assigned_to.email,
                    f"Assigned: Work Order #{wo.id}",
                    f"You have been assigned to work order #{wo.id}: {wo.title}",
                )
        except Exception:
            # fallback to synchronous send_mail if Celery not configured
            try:
                from django.core.mail import send_mail

                if wo.assigned_to and wo.assigned_to.email:
                    send_mail(
                        subject=f"Assigned: Work Order #{wo.id}",
                        message=f"You have been assigned to work order #{wo.id}: {wo.title}",
                        from_email=None,
                        recipient_list=[wo.assigned_to.email],
                        fail_silently=True,
                    )
            except Exception:
                pass

        return redirect("workorder_detail", wo_id=wo.id)

    # GET
    from django.contrib.auth import get_user_model

    User = get_user_model()
    users = User.objects.all()[:200]
    comments = wo.comments.all()
    attachments = wo.attachments.all()
    attachment_type_choices = WorkOrderAttachment.DOCUMENT_TYPE_CHOICES
    # provide logs ordered newest-first for the template
    logs = wo.logs.order_by("-created_at")
    # technicians for card-based assignment UI
    from .models import Technician
    try:
        technicians = Technician.objects.filter(active=True).order_by("name")
    except Exception:
        technicians = []
    # master methods for technician dropdown
    from .models import MasterItem
    try:
        # load all items in the 'method' category and build parent/child groups
        all_methods = list(
            MasterItem.objects.filter(category__iexact="method").order_by("order", "label")
        )
        roots = [m for m in all_methods if m.parent_id is None]
        method_groups = []
        for r in roots:
            children = [c for c in all_methods if c.parent_id == r.id]
            method_groups.append({"root": r, "children": children})

        # orphan children (parent not present in roots)
        parent_ids = {r.id for r in roots}
        method_singles = [m for m in all_methods if m.parent_id is not None and m.parent_id not in parent_ids]
    except Exception:
        method_groups = []
        method_singles = []
    # master RCA items for technician analysis dropdown (grouped)
    try:
        all_rca = list(
            MasterItem.objects.filter(category__iexact="rca").order_by("order", "label")
        )
        rca_roots = [m for m in all_rca if m.parent_id is None]
        # level-1 children of root → these become optgroup labels
        rca_groups = []
        for root in rca_roots:
            level1 = [m for m in all_rca if m.parent_id == root.id]
            for grp in level1:
                children = [m for m in all_rca if m.parent_id == grp.id]
                rca_groups.append({"group": grp, "children": children})
    except Exception:
        rca_groups = []
    # spare parts already used on this work order
    used_spare_parts = wo.used_spare_parts.select_related("spare_part").all()
    pending_procurement_count = sum(1 for u in used_spare_parts if u.is_out_of_stock)

    # Parse technician_report back into individual fields for prefill
    prefill_cause = ""
    prefill_analysis = ""
    prefill_actions = ""
    if wo.technician_report:
        import re as _re
        for line in wo.technician_report.split("\n\n"):
            line = line.strip()
            if line.startswith("สาเหตุ: "):
                prefill_cause = line[len("สาเหตุ: "):]
            elif line.startswith("วิเคราะห์สาเหตุ: "):
                prefill_analysis = line[len("วิเคราะห์สาเหตุ: "):]
            elif line.startswith("การดำเนินงาน: "):
                prefill_actions = line[len("การดำเนินงาน: "):]

    # Serialize existing spare parts as JSON for the picker prefill
    import json as _json
    prefill_spare_parts = []
    for usp in used_spare_parts:
        if usp.spare_part:
            prefill_spare_parts.append({
                "id": usp.spare_part.id,
                "code": usp.spare_part.code,
                "name": usp.spare_part.name,
                "stock": usp.spare_part.quantity,
                "qty": usp.quantity,
            })
        elif usp.custom_name:
            prefill_spare_parts.append({
                "id": "custom",
                "custom_name": usp.custom_name,
                "custom_unit": usp.custom_unit or "ชิ้น",
                "custom_unit_cost": float(usp.custom_unit_cost) if usp.custom_unit_cost else None,
                "qty": usp.quantity,
            })
    prefill_spare_parts_json = _json.dumps(prefill_spare_parts, ensure_ascii=False)
    # Build `workorderTypePills` html to show type labels consistent with KPI logic
    try:
        from django.utils.html import escape
        from django.utils.safestring import mark_safe
        from django.db.models import Q
        # Try to reuse the title->type keyword map from reports if available
        try:
            from cmms.views_reports import _TITLE_TYPE_Q as _GLOBAL_TITLE_Q
        except Exception:
            _GLOBAL_TITLE_Q = {}

        type_labels_list = []
        try:
            # Prefer explicit M2M links
            wo_types = list(wo.workorder_types.all())
            if wo_types:
                type_labels_list = [(t.label or "").strip() for t in wo_types if (t.label or "").strip()]
            else:
                # Legacy char field set and not 'other' → map to MasterItem label when possible
                if wo.workorder_type and wo.workorder_type != 'other':
                    mi = MasterItem.objects.filter(category__iexact='workorder_type', code=wo.workorder_type).first()
                    if mi and (mi.label or '').strip():
                        type_labels_list = [(mi.label or '').strip()]
                    else:
                        try:
                            type_labels_list = [wo.get_workorder_type_display()]
                        except Exception:
                            type_labels_list = []
                else:
                    # Title/label based matches (mirror calculate_kpis behavior)
                    matches = []
                    master_types = MasterItem.objects.filter(category='workorder_type', active=True).order_by('order', 'label')
                    single_qs = WorkOrder.objects.filter(id=wo.id)
                    for m in master_types:
                        label = (m.label or '').strip()
                        q_expr = None
                        if m.code and m.code in _GLOBAL_TITLE_Q:
                            q_expr = _GLOBAL_TITLE_Q[m.code]
                            if label:
                                q_expr = q_expr | Q(title__icontains=label)
                        else:
                            if label:
                                q_expr = Q(title__icontains=label)
                        if q_expr and single_qs.filter(q_expr).exists():
                            matches.append(label)
                    type_labels_list = matches
        except Exception:
            type_labels_list = []

        if type_labels_list:
            pills_html = ''.join(
                f'<span class="wo-badge badge-type" style="margin-right:0.35rem;"><i class="bi bi-tag-fill"></i> {escape(lbl)}</span>'
                for lbl in type_labels_list
            )
        else:
            pills_html = '<span class="text-muted" style="font-size:0.9rem;">— ไม่มีประเภทงาน —</span>'

        workorder_type_pills_html = mark_safe(pills_html)
    except Exception:
        from django.utils.safestring import mark_safe

        workorder_type_pills_html = mark_safe('<span class="text-muted" style="font-size:0.9rem;">— ไม่ทราบประเภทงาน —</span>')

    return render(
        request,
        "workorder/workorder_detail.html",
        {
            "wo": wo,
            "users": users,
            "technicians": technicians,
            "methods": None,
            "method_groups": method_groups,
            "method_singles": method_singles,
            "comments": comments,
            "attachments": attachments,
            "attachment_type_choices": attachment_type_choices,
            "logs": logs,
            "rca_groups": rca_groups,
            "used_spare_parts": used_spare_parts,
            "has_pending_procurement": pending_procurement_count > 0,
            "pending_procurement_count": pending_procurement_count,
            "prefill_cause": prefill_cause,
            "prefill_analysis": prefill_analysis,
            "prefill_actions": prefill_actions,
            "prefill_spare_parts_json": prefill_spare_parts_json,
            "workorderTypePills": workorder_type_pills_html,
            "wo_checklist": getattr(wo, "checklist", None),
            "wo_is_pm_or_calibration": bool(
                _resolve_workorder_type_codes(wo) & {"maintenance", "calibration"}
            ),
        },
    )


# สร้าง list สำหรับ dropdown
def customer_list(request):

    q = request.GET.get("q", "").strip()

    # support both singular and plural param names (forms/templates may vary)
    customer_units_selected = request.GET.getlist(
        "customer_units"
    ) or request.GET.getlist("customer_unit")
    customer_sections_selected = request.GET.getlist(
        "customer_section"
    ) or request.GET.getlist("customer_sections")
    customer_semi_departments_selected = request.GET.getlist(
        "customer_semi_department"
    ) or request.GET.getlist("customer_semi_departments")
    customer_departments_selected = request.GET.getlist(
        "customer_department"
    ) or request.GET.getlist("customer_departments")
    customer_divisions_selected = request.GET.getlist(
        "customer_division"
    ) or request.GET.getlist("customer_divisions")
    customer_customers_selected = request.GET.getlist(
        "customer_customer"
    ) or request.GET.getlist("customer_customers")

    all_customers = Customer_list.objects.all()
    if q:
        from django.db.models import Q

        all_customers = all_customers.filter(
            Q(customer_id__icontains=q)
            | Q(customer_units__icontains=q)
            | Q(customer_section__icontains=q)
            | Q(customer_semi_department__icontains=q)
            | Q(customer_department__icontains=q)
            | Q(customer_division__icontains=q)
            | Q(customer_customer__icontains=q)
            | Q(customer_name__icontains=q)
            | Q(customer_tel__icontains=q)
            | Q(customer_email__icontains=q)
        )

    # Apply multi-select filters only when there are non-empty values
    if customer_units_selected and any(val for val in customer_units_selected if val):
        all_customers = all_customers.filter(
            customer_units__in=[val for val in customer_units_selected if val]
        )
    if customer_sections_selected and any(
        val for val in customer_sections_selected if val
    ):
        all_customers = all_customers.filter(
            customer_section__in=[val for val in customer_sections_selected if val]
        )
    if customer_semi_departments_selected and any(
        val for val in customer_semi_departments_selected if val
    ):
        all_customers = all_customers.filter(
            customer_semi_department__in=[
                val for val in customer_semi_departments_selected if val
            ]
        )
    if customer_departments_selected and any(
        val for val in customer_departments_selected if val
    ):
        all_customers = all_customers.filter(
            customer_department__in=[
                val for val in customer_departments_selected if val
            ]
        )
    if customer_divisions_selected and any(
        val for val in customer_divisions_selected if val
    ):
        all_customers = all_customers.filter(
            customer_division__in=[val for val in customer_divisions_selected if val]
        )
    if customer_customers_selected and any(
        val for val in customer_customers_selected if val
    ):
        all_customers = all_customers.filter(
            customer_customer__in=[val for val in customer_customers_selected if val]
        )

    # Build lists for dropdown filters
    customer_units = (
        Customer_list.objects.values_list("customer_units", flat=True)
        .distinct()
        .order_by("customer_units")
    )
    customer_sections = (
        Customer_list.objects.values_list("customer_section", flat=True)
        .distinct()
        .order_by("customer_section")
    )
    customer_semi_departments = (
        Customer_list.objects.values_list("customer_semi_department", flat=True)
        .distinct()
        .order_by("customer_semi_department")
    )
    customer_departments = (
        Customer_list.objects.values_list("customer_department", flat=True)
        .distinct()
        .order_by("customer_department")
    )
    customer_divisions = (
        Customer_list.objects.values_list("customer_division", flat=True)
        .distinct()
        .order_by("customer_division")
    )
    customer_customers = (
        Customer_list.objects.values_list("customer_customer", flat=True)
        .distinct()
        .order_by("customer_customer")
    )

    return render(
        request,
        "customer/customer_list.html",
        {
            "all_customers": all_customers,
            "customer_units": customer_units,
            "customer_sections": customer_sections,
            "customer_semi_departments": customer_semi_departments,
            "customer_departments": customer_departments,
            "customer_divisions": customer_divisions,
            "customer_customers": customer_customers,
            "customer_units_selected": customer_units_selected,
            "customer_sections_selected": customer_sections_selected,
            "customer_departments_selected": customer_departments_selected,
            "customer_divisions_selected": customer_divisions_selected,
            "customer_customers_selected": customer_customers_selected,
            "q": q,
        },
    )


def add_customer(request):
    if request.method == "POST":
        customer_id = request.POST["customer_id"]
        customer_code = request.POST["customer_code"]
        customer_unit = request.POST["customer_unit"]
        customer_section = request.POST["customer_section"]
        customer_semi_department = request.POST["customer_semi_department"]
        customer_department = request.POST["customer_department"]
        customer_division = request.POST["customer_division"]
        customer_customer = request.POST["customer_customer"]
        customer_name = request.POST["customer_name"]
        customer_tel = request.POST["customer_tel"]
        customer_email = request.POST["customer_email"]
        customer_address = request.POST["customer_address"]
        customer_note = request.POST["customer_note"]

        # บันทึกข้อมูล
        customer_list = Customer_list.objects.create(
            customer_id=customer_id,
            customer_code=customer_code,
            customer_unit=customer_unit,
            customer_section=customer_section,
            customer_semi_department=customer_semi_department,
            customer_department=customer_department,
            customer_division=customer_division,
            customer_customer=customer_customer,
            customer_name=customer_name,
            customer_tel=customer_tel,
            customer_email=customer_email,
            customer_address=customer_address,
            customer_note=customer_note,
        )
        customer_list.save()
        messages.success(request, "บันทึกข้อมูลเรียบร้อยแล้ว")
        # เปลี่ยนเส้นทางไปหน้า customer_list
        return redirect("/customer_list")
    else:
        return render(request, "customer/add_customer.html")


def edit_customer(request, customer_list_id):
    if request.method == "POST":
        customer_list = Customer_list.objects.get(id=customer_list_id)
        customer_list.customer_id = request.POST["customer_id"]
        customer_list.customer_code = request.POST["customer_code"]
        customer_list.customer_unit = request.POST["customer_unit"]
        customer_list.customer_section = request.POST["customer_section"]
        customer_list.customer_semi_department = request.POST[
            "customer_semi_department"
        ]
        customer_list.customer_department = request.POST["customer_department"]
        customer_list.customer_division = request.POST["customer_division"]
        customer_list.customer_customer = request.POST["customer_customer"]
        customer_list.customer_name = request.POST["customer_name"]
        customer_list.customer_tel = request.POST["customer_tel"]
        customer_list.customer_email = request.POST["customer_email"]
        customer_list.customer_address = request.POST["customer_address"]
        customer_list.customer_note = request.POST["customer_note"]
        customer_list.save()
        messages.success(request, "แก้ไขข้อมูลเรียบร้อยแล้ว")
        return redirect("/customer_list")

    else:
        # ดึงข้อมูลเก่าขึ้นมาแสดง
        customer_list = Customer_list.objects.get(id=customer_list_id)
        return render(
            request, "customer/edit_customer.html", {"customer_list": customer_list}
        )


def delete_customer(request, customer_list_id):
    customer_list = Customer_list.objects.get(id=customer_list_id)
    customer_list.delete()
    messages.success(request, "ลบข้อมูลเรียบร้อยแล้ว")
    return redirect("/customer_list")


# อัพโหลดไฟล์ Excel
def import_customer_list(request):
    if request.method == "POST":
        excel_file = request.FILES["excel_file"]
        df = pd.read_excel(excel_file)

        for _, row in df.iterrows():
            Customer_list.objects.create(
                customer_id=row["customer_id"],
                customer_code=row["customer_code"],
                customer_unit=row["customer_unit"],
                customer_section=row["customer_section"],
                customer_semi_department=row["customer_semi_department"],
                customer_department=row["customer_department"],
                customer_division=row["customer_division"],
                customer_customer=row["customer_customer"],
                customer_name=row["customer_name"],
                customer_tel=row["customer_tel"],
                customer_email=row["customer_email"],
                customer_address=row["customer_address"],
                customer_note=row["customer_note"],
            )
        messages.success(request, "นำเข้าข้อมูลเรียบร้อยแล้ว")
        return redirect("customer_list")
    return render(request, "customer/import_customer_list.html")


# ดาวน์โหลดไฟล์ Excel
def export_customer_list(request):
    q = request.GET.get("q", "").strip()
    customer_units_selected = request.GET.getlist(
        "customer_units"
    ) or request.GET.getlist("customer_unit")
    customer_sections_selected = request.GET.getlist(
        "customer_section"
    ) or request.GET.getlist("customer_sections")
    customer_semi_departments_selected = request.GET.getlist(
        "customer_semi_department"
    ) or request.GET.getlist("customer_semi_departments")
    customer_departments_selected = request.GET.getlist(
        "customer_department"
    ) or request.GET.getlist("customer_departments")
    customer_divisions_selected = request.GET.getlist(
        "customer_division"
    ) or request.GET.getlist("customer_divisions")
    customer_customers_selected = request.GET.getlist(
        "customer_customer"
    ) or request.GET.getlist("customer_customers")

    all_customers = Customer_list.objects.all()
    if q:
        from django.db.models import Q

        all_customers = all_customers.filter(
            Q(customer_id__icontains=q)
            | Q(customer_units__icontains=q)
            | Q(customer_section__icontains=q)
            | Q(customer_semi_department__icontains=q)
            | Q(customer_department__icontains=q)
            | Q(customer_division__icontains=q)
            | Q(customer_customer__icontains=q)
            | Q(customer_name__icontains=q)
            | Q(customer_tel__icontains=q)
            | Q(customer_email__icontains=q)
        )

    # Apply multi-select filters only when there are non-empty values
    if customer_units_selected and any(val for val in customer_units_selected if val):
        all_customers = all_customers.filter(
            customer_units__in=[val for val in customer_units_selected if val]
        )
    if customer_sections_selected and any(
        val for val in customer_sections_selected if val
    ):
        all_customers = all_customers.filter(
            customer_section__in=[val for val in customer_sections_selected if val]
        )
    if customer_semi_departments_selected and any(
        val for val in customer_semi_departments_selected if val
    ):
        all_customers = all_customers.filter(
            customer_semi_department__in=[
                val for val in customer_semi_departments_selected if val
            ]
        )
    if customer_departments_selected and any(
        val for val in customer_departments_selected if val
    ):
        all_customers = all_customers.filter(
            customer_department__in=[
                val for val in customer_departments_selected if val
            ]
        )
    if customer_divisions_selected and any(
        val for val in customer_divisions_selected if val
    ):
        all_customers = all_customers.filter(
            customer_division__in=[val for val in customer_divisions_selected if val]
        )
    if customer_customers_selected and any(
        val for val in customer_customers_selected if val
    ):
        all_customers = all_customers.filter(
            customer_customer__in=[val for val in customer_customers_selected if val]
        )
        # Export to Excel
        if pd is None:
            return HttpResponse(
                "Server error: pandas not installed (required to export Excel).",
                status=500,
            )

        customer_list = all_customers.values()
        df = pd.DataFrame(customer_list)
        response = HttpResponse(
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        response["Content-Disposition"] = "attachment; filename=customer_list.xlsx"
        with pd.ExcelWriter(response, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Customer_list")
        return response


# Minimal views referenced by urls.py but previously missing
def master_data(request):
    # Build a dynamic menu for master data. Each item may provide a url_name (Django named route)
    # and/or a fallback path. This keeps template simple and robust if some routes are missing.
    menu = [
        {
            "label": "Equipment Images",
            "url_name": "equipment_images",
            "path": "/equipment-images/",
        },
        {
            "label": "Equipment Master List",
            "url_name": "equipment_list",
            "path": "/equipment_list/",
        },
        {
            "label": "Customer Structure",
            "url_name": "customer_list",
            "path": "/customer_list",
        },
        {"label": "User Appointments", "path": "/user_appointments"},
        {"label": "Service Contracts", "path": "/service_contracts"},
        {"label": "Suppliers", "path": "/suppliers"},
        {"label": "Technicians", "path": "/technicians"},
        {"label": "Equipment Types", "path": "/equipment_types"},
        {"label": "Equipment Statuses", "path": "/equipment_statuses"},
        {"label": "Fault Types", "path": "/fault_types"},
        {"label": "Repair Methods", "path": "/repair_methods"},
        {"label": "Parts", "path": "/parts"},
        {"label": "Manufacturers", "path": "/manufacturers"},
        {"label": "Locations", "path": "/locations"},
        {"label": "Departments", "path": "/departments"},
        {"label": "Vendors", "path": "/vendors"},
        {"label": "Priorities", "path": "/priorities"},
        {"label": "Purchase Orders", "path": "/purchase_orders"},
        {"label": "Warranty Types", "path": "/warranty_types"},
        {"label": "Currencies", "path": "/currencies"},
        {"label": "Users", "path": "/users"},
        {"label": "Roles", "path": "/roles"},
        {"label": "Permissions", "path": "/permissions"},
    ]
    return render(request, "master/master_data.html", {"menu": menu})


@staff_member_required
def master_tree_manage(request, category=None):
    # reuse master_list's tree building logic but render a management UI
    categories = list(
        MasterItem.objects.values_list("category", flat=True)
        .distinct()
        .order_by("category")
    )
    tree_roots = None
    active_category = None
    if category:
        active_category = category
        nodes = list(
            MasterItem.objects.filter(category=category).order_by("order", "label")
        )
        node_map = {n.id: {"item": n, "children": []} for n in nodes}
        roots = []
        for n in nodes:
            if n.parent_id and n.parent_id in node_map:
                node_map[n.parent_id]["children"].append(node_map[n.id])
            else:
                roots.append(node_map[n.id])
        tree_roots = roots
    return render(
        request,
        "master/tree_manage.html",
        {
            "categories": categories,
            "active_category": active_category,
            "tree_roots": tree_roots,
        },
    )


@staff_member_required
@require_http_methods(["POST"])
def master_api_create(request):
    # expects: category, code, label, description, active (optional), order (optional), parent_id (optional)
    data = request.POST or request.POST
    # diagnostic: capture incoming parent and posted keys
    try:
        import logging

        logger = logging.getLogger(__name__)
        logger.info(
            "master_api_create called: session=%s keys=%s",
            getattr(request.session, "session_key", None),
            list(data.keys()),
        )
        logger.debug(
            "master_api_create POST parent_id=%s csrf_in_post=%s csrf_cookie=%s",
            data.get("parent_id"),
            ("csrfmiddlewaretoken" in data),
            bool(request.COOKIES.get("csrftoken")),
        )
    except Exception:
        pass
    category = (data.get("category") or "").strip()
    code = (data.get("code") or "").strip()
    label = (data.get("label") or "").strip()
    description = data.get("description") or ""
    active = data.get("active", "true") in ("1", "true", "True")
    order = data.get("order") or 0
    parent_id = data.get("parent_id")
    if not category or not label:
        return JsonResponse(
            {"ok": False, "error": "category and label required"}, status=400
        )
    # validate unique (category, code) when code provided
    if code:
        if MasterItem.objects.filter(category=category, code=code).exists():
            return JsonResponse(
                {"ok": False, "error": "Duplicate code in this category"}, status=400
            )
    item = MasterItem.objects.create(
        category=category,
        code=code,
        label=label,
        description=description,
        active=active,
        order=order,
    )
    if parent_id:
        try:
            parent = MasterItem.objects.get(pk=int(parent_id))
            # allow cross-category parenting; do not change item's category
            item.parent = parent
            item.save()
        except Exception:
            pass
    # push undo snapshot so this create can be undone (delete)
    try:
        _push_undo_snapshot(request, action="create", payload={"id": item.pk})
    except Exception:
        pass
    # debug: include parent info in response so client can confirm what was saved
    try:
        import logging

        logger = logging.getLogger(__name__)
        logger.info(
            "master_api_create: created id=%s parent_id=%s", item.pk, item.parent_id
        )
    except Exception:
        pass
    parent_label = ""
    try:
        if item.parent:
            parent_label = f"{item.parent.label} ({item.parent.code or ''})"
    except Exception:
        parent_label = ""
    # include submitted parent and posted keys for diagnostics
    submitted_parent = data.get("parent_id")
    posted_keys = list(data.keys())
    return JsonResponse(
        {
            "ok": True,
            "id": item.pk,
            "label": item.label,
            "code": item.code,
            "parent_id": item.parent_id,
            "parent_label": parent_label,
            "submitted_parent": submitted_parent,
            "posted_keys": posted_keys,
        }
    )


@staff_member_required
@require_http_methods(["POST"])
def master_api_update(request, pk):
    item = get_object_or_404(MasterItem, pk=pk)
    data = request.POST
    # diagnostic: log incoming payload and session
    try:
        import logging

        logger = logging.getLogger(__name__)
        logger.info(
            "master_api_update called id=%s session=%s keys=%s",
            pk,
            getattr(request.session, "session_key", None),
            list(data.keys()),
        )
        logger.debug(
            "master_api_update POST parent_id=%s csrf_in_post=%s csrf_cookie=%s",
            data.get("parent_id"),
            ("csrfmiddlewaretoken" in data),
            bool(request.COOKIES.get("csrftoken")),
        )
    except Exception:
        pass
    # capture before state for undo
    before = {
        "category": item.category,
        "code": item.code,
        "label": item.label,
        "description": item.description,
        "active": bool(item.active),
        "order": item.order,
        "parent_id": item.parent_id,
    }
    try:
        _push_undo_snapshot(
            request, action="update", payload={"id": item.pk, "before": before}
        )
    except Exception:
        pass

    item.category = (data.get("category") or item.category).strip()
    item.code = (data.get("code") or item.code).strip()
    item.label = (data.get("label") or item.label).strip()
    item.description = data.get("description") or item.description
    item.active = data.get("active", "true") in ("1", "true", "True")
    try:
        item.order = int(data.get("order"))
    except Exception:
        pass
    parent_id = data.get("parent_id")
    # keep track of what was submitted
    submitted_parent = parent_id
    if parent_id:
        try:
            parent = MasterItem.objects.get(pk=int(parent_id))
            # allow cross-category parenting; do not change item's category
            item.parent = parent
        except Exception:
            item.parent = None
    else:
        item.parent = None
    # uniqueness check for code within category (exclude self)
    if item.code:
        if (
            MasterItem.objects.filter(category=item.category, code=item.code)
            .exclude(pk=item.pk)
            .exists()
        ):
            return JsonResponse(
                {"ok": False, "error": "Duplicate code in this category"}, status=400
            )
    item.save()
    # debug: log and return parent info so client can detect if parent was cleared
    try:
        import logging

        logger = logging.getLogger(__name__)
        logger.info(
            "master_api_update: updated id=%s parent_id=%s", item.pk, item.parent_id
        )
    except Exception:
        pass
    parent_label = ""
    try:
        if item.parent:
            parent_label = f"{item.parent.label} ({item.parent.code or ''})"
    except Exception:
        parent_label = ""
    # echo submitted parent and posted keys for debugging
    posted_keys = list(data.keys())
    return JsonResponse(
        {
            "ok": True,
            "parent_id": item.parent_id,
            "parent_label": parent_label,
            "submitted_parent": submitted_parent,
            "posted_keys": posted_keys,
        }
    )


@staff_member_required
@require_http_methods(["POST"])
def master_api_delete(request, pk):
    item = get_object_or_404(MasterItem, pk=pk)
    # capture before state so deletion can be undone
    before = {
        "category": item.category,
        "code": item.code,
        "label": item.label,
        "description": item.description,
        "active": bool(item.active),
        "order": item.order,
        "parent_id": item.parent_id,
    }
    try:
        _push_undo_snapshot(
            request, action="delete", payload={"id": item.pk, "before": before}
        )
    except Exception:
        pass
    item.delete()
    return JsonResponse({"ok": True})


@staff_member_required
@require_http_methods(["GET"])
def master_api_get(request, pk):
    item = get_object_or_404(MasterItem, pk=pk)
    parent_label = ""
    if item.parent:
        parent_label = f"{item.parent.label} ({item.parent.code or ''})"
    return JsonResponse(
        {
            "ok": True,
            "id": item.pk,
            "category": item.category,
            "code": item.code or "",
            "label": item.label,
            "description": item.description or "",
            "order": item.order or 0,
            "active": bool(item.active),
            "parent_id": item.parent_id,
            "parent_label": parent_label,
        }
    )


@staff_member_required
def master_api_parents(request):
    """Return candidates for parent selection: ?category=...&q=term"""
    cat = (request.GET.get("category") or "").strip()
    q = (request.GET.get("q") or "").strip()
    qs = MasterItem.objects.all()
    if cat:
        qs = qs.filter(category=cat)
    if q:
        from django.db.models import Q

        qs = qs.filter(Q(label__icontains=q) | Q(code__icontains=q))
    out = [
        {"id": m.pk, "text": f"{m.label} ({m.code or ''})"}
        for m in qs.order_by("label")[:50]
    ]
    return JsonResponse({"results": out})


@staff_member_required
@require_http_methods(["POST"])
def master_api_reorder(request):
    """Accept JSON payload: { parent_id: <id|null>, ordered_ids: [id1,id2,...] }
    Update each item's parent and order (order = index*10) within that parent/category.
    """
    try:
        import json

        body = request.body.decode("utf-8") if request.body else ""
        data = json.loads(body) if body else {}
    except Exception:
        return JsonResponse({"ok": False, "error": "invalid JSON"}, status=400)

    ordered = data.get("ordered_ids") or []
    parent_id = data.get("parent_id")
    try:
        parent = MasterItem.objects.get(pk=int(parent_id)) if parent_id else None
    except Exception:
        parent = None

    # Update items: set parent and order (step 10). We do NOT change item.category here.
    # capture snapshot before changes for undo: include previous parent_id and order for each affected item
    before_order = []
    try:
        for sid in ordered:
            try:
                mi = MasterItem.objects.get(pk=int(sid))
                before_order.append(
                    {"id": mi.pk, "parent_id": mi.parent_id, "order": mi.order}
                )
            except Exception:
                continue
    except Exception:
        before_order = []

    try:
        _push_undo_snapshot(
            request,
            action="reorder",
            payload={
                "parent_id": parent_id,
                "ordered_ids": ordered,
                "before_order": before_order,
            },
        )
    except Exception:
        # best-effort only; do not fail operation on snapshot error
        pass

    for idx, sid in enumerate(ordered):
        try:
            mi = MasterItem.objects.get(pk=int(sid))
            mi.parent = parent
            mi.order = idx * 10
            mi.save()
        except Exception:
            # skip invalid ids
            continue

    return JsonResponse({"ok": True})


### Session-backed undo/redo helpers and endpoints -------------------------
def _get_undo_stack(session):
    return session.get("master_undo_stack", [])


def _set_undo_stack(session, stack):
    session["master_undo_stack"] = stack
    session.modified = True


def _get_redo_stack(session):
    return session.get("master_redo_stack", [])


def _set_redo_stack(session, stack):
    session["master_redo_stack"] = stack
    session.modified = True


def _push_undo_snapshot(request, action, payload):
    """Push a snapshot describing a mutating action so frontend can request an undo.

    Snapshot format: { action: 'create'|'update'|'delete'|'reorder', payload: {...} }
    For 'create' payload: { id: <created_id> }
    For 'update' payload: { id: <id>, before: { ...fields... } }
    For 'delete' payload: { id: <id>, before: { ...fields... } }
    For 'reorder' payload: { parent_id: <id|null>, ordered_ids: [...] , before_order: [{id, parent_id, order}, ...] }
    """
    # best-effort: record minimal snapshot data
    stack = _get_undo_stack(request.session)
    stack.append({"action": action, "payload": payload})
    # cap stack size to avoid unbounded growth
    if len(stack) > 20:
        stack = stack[-20:]
    _set_undo_stack(request.session, stack)
    # clear redo on new action
    _set_redo_stack(request.session, [])


@staff_member_required
@require_http_methods(["POST"])
def master_api_undo(request):
    """Pop last action from session undo stack and attempt to revert it.

    Returns JSON { ok: True, restored: <info> } or error.
    This is intentionally conservative: it attempts to reverse the last action
    by calling the appropriate DB operations. Complex conflicts (concurrent
    edits by other users) may prevent a perfect revert.
    """
    stack = _get_undo_stack(request.session)
    if not stack:
        return JsonResponse({"ok": False, "error": "nothing to undo"}, status=400)
    snap = stack.pop()
    # save updated stack
    _set_undo_stack(request.session, stack)
    action = snap.get("action")
    payload = snap.get("payload") or {}
    try:
        if action == "create":
            # created item -> delete it
            cid = payload.get("id")
            if cid:
                MasterItem.objects.filter(pk=cid).delete()
                # push to redo stack the inverse (recreate)
                rs = _get_redo_stack(request.session)
                rs.append({"action": "create", "payload": {"id": cid}})
                _set_redo_stack(request.session, rs)
                return JsonResponse(
                    {"ok": True, "restored": {"action": "delete", "id": cid}}
                )
            return JsonResponse({"ok": False, "error": "invalid snapshot"}, status=400)

        if action == "delete":
            # delete snapshot must carry 'before' state to recreate
            before = payload.get("before") or {}
            if not before:
                return JsonResponse(
                    {"ok": False, "error": "missing before state for delete"},
                    status=400,
                )
            # recreate object
            m = MasterItem.objects.create(
                category=before.get("category") or "",
                code=before.get("code") or "",
                label=before.get("label") or "",
                description=before.get("description") or "",
                active=bool(before.get("active", True)),
                order=int(before.get("order") or 0),
            )
            # try to restore parent if present
            pid = before.get("parent_id")
            if pid:
                try:
                    p = MasterItem.objects.get(pk=int(pid))
                    m.parent = p
                    m.save()
                except Exception:
                    pass
            # push redo snapshot
            rs = _get_redo_stack(request.session)
            rs.append({"action": "delete", "payload": {"id": m.pk}})
            _set_redo_stack(request.session, rs)
            return JsonResponse(
                {"ok": True, "restored": {"action": "create", "id": m.pk}}
            )

        if action == "update":
            before = payload.get("before") or {}
            mid = payload.get("id")
            if not mid or not before:
                return JsonResponse(
                    {"ok": False, "error": "invalid snapshot"}, status=400
                )
            mi = MasterItem.objects.filter(pk=mid).first()
            if not mi:
                return JsonResponse({"ok": False, "error": "item missing"}, status=404)
            # store current state for redo
            current = {
                "category": mi.category,
                "code": mi.code,
                "label": mi.label,
                "description": mi.description,
                "active": bool(mi.active),
                "order": mi.order,
                "parent_id": mi.parent_id,
            }
            # restore fields
            mi.category = before.get("category") or mi.category
            mi.code = before.get("code") or mi.code
            mi.label = before.get("label") or mi.label
            mi.description = before.get("description") or mi.description
            mi.active = bool(before.get("active", mi.active))
            try:
                mi.order = int(before.get("order") or mi.order)
            except Exception:
                pass
            pid = before.get("parent_id")
            if pid:
                try:
                    mi.parent = MasterItem.objects.get(pk=int(pid))
                except Exception:
                    mi.parent = None
            else:
                mi.parent = None
            mi.save()
            # push redo snapshot
            rs = _get_redo_stack(request.session)
            rs.append({"action": "update", "payload": {"id": mid, "before": current}})
            _set_redo_stack(request.session, rs)
            return JsonResponse(
                {"ok": True, "restored": {"action": "update", "id": mid}}
            )

        if action == "reorder":
            # payload should contain 'before_order' to fully restore; if missing, we can try
            before_order = payload.get("before_order") or []
            # perform restore per entry
            for entry in before_order:
                try:
                    mi = MasterItem.objects.get(pk=int(entry.get("id")))
                    mi.parent_id = entry.get("parent_id")
                    try:
                        mi.order = int(entry.get("order") or mi.order)
                    except Exception:
                        pass
                    mi.save()
                except Exception:
                    continue
            # push redo snapshot to reapply the reorder we just undid
            rs = _get_redo_stack(request.session)
            rs.append({"action": "reorder", "payload": payload})
            _set_redo_stack(request.session, rs)
            return JsonResponse({"ok": True, "restored": {"action": "reorder"}})

    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=500)

    return JsonResponse({"ok": False, "error": "unsupported action"}, status=400)


@staff_member_required
@require_http_methods(["POST"])
def master_api_redo(request):
    """Re-apply last undone action from redo stack."""
    rs = _get_redo_stack(request.session)
    if not rs:
        return JsonResponse({"ok": False, "error": "nothing to redo"}, status=400)
    snap = rs.pop()
    _set_redo_stack(request.session, rs)
    action = snap.get("action")
    payload = snap.get("payload") or {}
    try:
        if action == "create":
            # recreate previously-created item? We don't have full data; best-effort: nothing to do
            return JsonResponse(
                {"ok": False, "error": "cannot redo create without snapshot data"},
                status=400,
            )

        if action == "delete":
            # delete the item again
            cid = payload.get("id")
            if cid:
                MasterItem.objects.filter(pk=cid).delete()
                # push to undo
                us = _get_undo_stack(request.session)
                us.append({"action": "delete", "payload": {"id": cid}})
                _set_undo_stack(request.session, us)
                return JsonResponse({"ok": True})
            return JsonResponse({"ok": False, "error": "invalid snapshot"}, status=400)

        if action == "update":
            # reapply update from payload.before (which was previous state). For redo we expect payload.before to be the state to set.
            mid = payload.get("id")
            before = payload.get("before") or {}
            if not mid or not before:
                return JsonResponse(
                    {"ok": False, "error": "invalid snapshot"}, status=400
                )
            mi = MasterItem.objects.filter(pk=mid).first()
            if not mi:
                return JsonResponse({"ok": False, "error": "item missing"}, status=404)
            # store current to undo
            current = {
                "category": mi.category,
                "code": mi.code,
                "label": mi.label,
                "description": mi.description,
                "active": bool(mi.active),
                "order": mi.order,
                "parent_id": mi.parent_id,
            }
            # apply redo state
            mi.category = before.get("category") or mi.category
            mi.code = before.get("code") or mi.code
            mi.label = before.get("label") or mi.label
            mi.description = before.get("description") or mi.description
            mi.active = bool(before.get("active", mi.active))
            try:
                mi.order = int(before.get("order") or mi.order)
            except Exception:
                pass
            pid = before.get("parent_id")
            if pid:
                try:
                    mi.parent = MasterItem.objects.get(pk=int(pid))
                except Exception:
                    mi.parent = None
            else:
                mi.parent = None
            mi.save()
            us = _get_undo_stack(request.session)
            us.append({"action": "update", "payload": {"id": mid, "before": current}})
            _set_undo_stack(request.session, us)
            return JsonResponse({"ok": True})

        if action == "reorder":
            ordered = payload.get("ordered_ids") or []
            parent_id = payload.get("parent_id")
            try:
                parent = (
                    MasterItem.objects.get(pk=int(parent_id)) if parent_id else None
                )
            except Exception:
                parent = None
            # capture before-order for undo
            before_order = payload.get("before_order") or []
            # apply reorder
            for idx, sid in enumerate(ordered):
                try:
                    mi = MasterItem.objects.get(pk=int(sid))
                    mi.parent = parent
                    mi.order = idx * 10
                    mi.save()
                except Exception:
                    continue
            us = _get_undo_stack(request.session)
            us.append(
                {
                    "action": "reorder",
                    "payload": {
                        "parent_id": parent_id,
                        "ordered_ids": ordered,
                        "before_order": before_order,
                    },
                }
            )
            _set_undo_stack(request.session, us)
            return JsonResponse({"ok": True})
    except Exception as e:
        return JsonResponse({"ok": False, "error": str(e)}, status=500)

    return JsonResponse({"ok": False, "error": "unsupported redo action"}, status=400)


def reports(request):
    # Placeholder reports page
    return render(request, "misc/reports.html")


def about(request):
    # Simple about page
    return render(request, "misc/about.html")


def create_account(request):
    """Public registration page — creates a Django user + Profile."""
    from django.contrib.auth import get_user_model
    from django.contrib.auth.password_validation import validate_password

    from .models import Profile

    User = get_user_model()
    department_options = MasterItem.objects.filter(
        category="customers", active=True
    ).order_by("order", "label").values_list("label", flat=True)

    if request.method == "POST":
        data = request.POST
        errors = {}
        form_values = {
            "username": data.get("username", "").strip(),
            "first_name": data.get("first_name", "").strip(),
            "last_name": data.get("last_name", "").strip(),
            "email": data.get("email", "").strip(),
            "telephone": data.get("telephone", "").strip(),
            "department": data.get("department", "").strip(),
            "position": data.get("position", "").strip(),
        }

        username = form_values["username"]
        password = data.get("password", "")
        password2 = data.get("password2", "")

        if not username:
            errors["username"] = "กรุณากรอกชื่อผู้ใช้"
        elif User.objects.filter(username=username).exists():
            errors["username"] = "ชื่อผู้ใช้นี้ถูกใช้งานแล้ว"

        if not password:
            errors["password"] = "กรุณากรอกรหัสผ่าน"
        else:
            try:
                validate_password(password)
            except Exception as e:
                errors["password"] = " ".join(e.messages)

        if password and password != password2:
            errors["password2"] = "รหัสผ่านไม่ตรงกัน"

        if not form_values["first_name"]:
            errors["first_name"] = "กรุณากรอกชื่อ"
        if not form_values["last_name"]:
            errors["last_name"] = "กรุณากรอกนามสกุล"

        if errors:
            return render(request, "create_account.html", {
                "errors": errors,
                "form_values": form_values,
                "department_options": department_options,
            })

        user = User.objects.create_user(
            username=username,
            password=password,
            first_name=form_values["first_name"],
            last_name=form_values["last_name"],
            email=form_values["email"],
        )
        # New self-registered accounts are inactive until an admin/staff
        # reviews and approves them.
        user.is_active = False
        user.save(update_fields=["is_active"])

        # Profile is auto-created by post_save signal; update extra fields
        profile = user.profile
        profile.telephone = form_values["telephone"]
        profile.department = form_values["department"]
        profile.position = form_values["position"]
        profile.approval_status = Profile.APPROVAL_PENDING
        profile.save()

        messages.success(
            request,
            "สร้างบัญชีสำเร็จ บัญชีของคุณอยู่ระหว่างการตรวจสอบโดยผู้ดูแลระบบ "
            "กรุณารอการอนุมัติก่อนเข้าสู่ระบบ",
        )
        return redirect("login")

    return render(
        request,
        "create_account.html",
        {"department_options": department_options},
    )


def _can_approve_accounts(user):
    return user.is_authenticated and (
        user.is_superuser or user.is_staff or user.has_perm("cmms.can_approve_accounts")
    )


@user_passes_test(_can_approve_accounts)
def account_approvals(request):
    """List new-account registrations for staff to review and approve/reject."""
    from .models import Profile

    status = request.GET.get("status", "pending")
    qs = Profile.objects.select_related("user").order_by("-created_at")
    if status in dict(Profile.APPROVAL_STATUS_CHOICES):
        qs = qs.filter(approval_status=status)

    pending_count = Profile.objects.filter(
        approval_status=Profile.APPROVAL_PENDING
    ).count()

    from django.contrib.auth.models import Group

    return render(
        request,
        "account_approvals.html",
        {
            "profiles": qs,
            "status": status,
            "pending_count": pending_count,
            "groups": Group.objects.all(),
        },
    )


@user_passes_test(_can_approve_accounts)
def account_approve(request, user_id):
    from .models import Profile

    if request.method != "POST":
        return redirect("account_approvals")

    from django.contrib.auth import get_user_model
    from django.contrib.auth.models import Group

    User = get_user_model()
    target_user = get_object_or_404(User, id=user_id)
    profile, _ = Profile.objects.get_or_create(user=target_user)

    target_user.is_active = True
    target_user.save(update_fields=["is_active"])

    group_ids = request.POST.getlist("groups")
    if group_ids:
        target_user.groups.set(Group.objects.filter(id__in=group_ids))

    profile.approval_status = Profile.APPROVAL_APPROVED
    profile.approved_by = request.user
    profile.approved_at = timezone.now()
    profile.rejection_reason = ""
    profile.save()

    messages.success(request, f"อนุมัติบัญชี {target_user.username} เรียบร้อยแล้ว")
    return redirect("account_approvals")


@user_passes_test(_can_approve_accounts)
def account_reject(request, user_id):
    from .models import Profile

    if request.method != "POST":
        return redirect("account_approvals")

    from django.contrib.auth import get_user_model

    User = get_user_model()
    target_user = get_object_or_404(User, id=user_id)
    profile, _ = Profile.objects.get_or_create(user=target_user)

    target_user.is_active = False
    target_user.save(update_fields=["is_active"])

    profile.approval_status = Profile.APPROVAL_REJECTED
    profile.approved_by = request.user
    profile.approved_at = timezone.now()
    profile.rejection_reason = request.POST.get("reason", "").strip()
    profile.save()

    messages.warning(request, f"ปฏิเสธคำขอสมัครบัญชี {target_user.username} แล้ว")
    return redirect("account_approvals")


def _can_manage_permissions(user):
    return user.is_authenticated and (user.is_superuser or user.is_staff)


@user_passes_test(_can_manage_permissions)
def user_management(request):
    """List all users so staff can review/change each user's groups (data access) and active status."""
    from django.contrib.auth import get_user_model
    from django.contrib.auth.models import Group
    from django.db.models import Q

    User = get_user_model()

    q = request.GET.get("q", "").strip()
    users = (
        User.objects.select_related("profile")
        .prefetch_related("groups")
        .order_by("username")
    )
    if q:
        users = users.filter(
            Q(username__icontains=q)
            | Q(first_name__icontains=q)
            | Q(last_name__icontains=q)
            | Q(email__icontains=q)
        )

    groups = Group.objects.prefetch_related("permissions__content_type").order_by("name")

    return render(
        request,
        "user_management.html",
        {
            "users": users,
            "groups": groups,
            "q": q,
        },
    )


@user_passes_test(_can_manage_permissions)
def user_update_access(request, user_id):
    """Update a single user's group memberships (data access/edit rights) and active status."""
    from django.contrib.auth import get_user_model
    from django.contrib.auth.models import Group

    from .models import Profile

    if request.method != "POST":
        return redirect("user_management")

    User = get_user_model()
    target_user = get_object_or_404(User, id=user_id)

    if target_user.is_superuser and not request.user.is_superuser:
        messages.error(request, "คุณไม่มีสิทธิ์แก้ไขบัญชีผู้ดูแลระบบสูงสุด (superuser)")
        return redirect("user_management")

    if target_user.id == request.user.id and "is_active" not in request.POST:
        messages.error(request, "ไม่สามารถระงับการใช้งานบัญชีของตัวเองได้")
        return redirect("user_management")

    group_ids = request.POST.getlist("groups")
    target_user.groups.set(Group.objects.filter(id__in=group_ids))

    was_active = target_user.is_active
    target_user.is_active = "is_active" in request.POST

    # Only a superuser can grant/revoke staff (Django admin) access.
    if request.user.is_superuser and target_user.id != request.user.id:
        target_user.is_staff = "is_staff" in request.POST

    target_user.save()

    profile, _ = Profile.objects.get_or_create(user=target_user)
    if target_user.is_active and not was_active:
        profile.approval_status = Profile.APPROVAL_APPROVED
        profile.approved_by = request.user
        profile.approved_at = timezone.now()
        profile.save()

    messages.success(request, f"บันทึกสิทธิ์การเข้าถึงของ {target_user.username} เรียบร้อยแล้ว")
    return redirect("user_management")


@login_required
def profile(request):
    """Simple user profile page used by the navbar link. Shows basic user info and groups."""
    user = request.user
    groups = user.groups.all() if user.is_authenticated else []
    profile = getattr(user, 'profile', None) if user.is_authenticated else None
    # Provide master_customers (with ancestor chains) so templates can show parent info
    try:
        qs = list(MasterItem.objects.filter(category="customers", active=True).order_by("order", "label"))
        if qs:
            label_map = {m.id: m.label for m in qs}
            parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}

            def build_ancestors(mid):
                chain = []
                seen = set()
                cur = parent_map.get(mid)
                while cur and cur not in seen:
                    seen.add(cur)
                    chain.insert(0, label_map.get(cur, ""))
                    cur = parent_map.get(cur)
                return " → ".join(chain) if chain else ""

            master_customers = [{"id": m.id, "label": m.label, "ancestors": build_ancestors(m.id)} for m in qs]
        else:
            master_customers = []
    except Exception:
        master_customers = []

    return render(request, "profile.html", {"user": user, "groups": groups, "profile": profile, "master_customers": master_customers})


@login_required
def edit_profile(request):
    user = request.user
    if request.method == "POST":
        form = ProfileForm(request.POST, instance=user)
        if form.is_valid():
            form.save()
            
            # Handle groups assignment
            if user.is_staff or user.is_superuser:
                selected_group_ids = request.POST.getlist('groups')
                user.groups.clear()
                for group_id in selected_group_ids:
                    try:
                        from django.contrib.auth.models import Group
                        group = Group.objects.get(id=int(group_id))
                        user.groups.add(group)
                    except (Group.DoesNotExist, ValueError):
                        pass
            
            messages.success(request, "บันทึกโปรไฟล์เรียบร้อยแล้ว")
            return redirect("profile")
    else:
        form = ProfileForm(instance=user)

    # Get master_customers with ancestor chains
    try:
        qs = list(MasterItem.objects.filter(category="customers", active=True).order_by("order", "label"))
        if qs:
            label_map = {m.id: m.label for m in qs}
            parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs}
            
            def build_ancestors(mid):
                chain = []
                seen = set()
                cur = parent_map.get(mid)
                while cur and cur not in seen:
                    seen.add(cur)
                    chain.insert(0, label_map.get(cur, ""))
                    cur = parent_map.get(cur)
                return " → ".join(chain) if chain else ""
            
            master_customers = [{"id": m.id, "label": m.label, "ancestors": build_ancestors(m.id)} for m in qs]
        else:
            master_customers = []
    except Exception:
        master_customers = []

    # Get all groups if user is staff
    from django.contrib.auth.models import Group
    groups = Group.objects.all() if (user.is_staff or user.is_superuser) else []

    # Determine current selections
    current_department = request.POST.get("department") if request.method == "POST" else (getattr(user, "profile", None) and user.profile.department or "")

    context = {
        "form": form,
        "master_customers": master_customers,
        "groups": groups,
        "form_values": {"department": current_department},
    }
    return render(request, "profile_edit.html", context)


@login_required
def change_password(request):
    if request.method == "POST":
        form = PasswordChangeForm(user=request.user, data=request.POST)
        if form.is_valid():
            user = form.save()
            update_session_auth_hash(request, user)  # keep user logged in
            messages.success(request, "เปลี่ยนรหัสผ่านเรียบร้อยแล้ว")
            return redirect("profile")
    else:
        form = PasswordChangeForm(user=request.user)
    return render(request, "change_password.html", {"form": form})


# --- Multi-account support -------------------------------------------------
from django.contrib.auth import authenticate, login


def _get_alt_accounts(session):
    return session.get("alt_accounts", [])


def _set_alt_accounts(session, accounts):
    session["alt_accounts"] = accounts
    session.modified = True


@login_required
def multi_accounts(request):
    """Show list of stored alternate accounts (usernames) for quick switching."""
    accounts = _get_alt_accounts(request.session)
    return render(request, "accounts/multi_accounts.html", {"accounts": accounts})


@login_required
def add_account(request):
    """Authenticate credentials (username+password) and store minimal info in session for switching.
    Expects POST with username and password. Does not change current user.
    """
    import logging

    logger = logging.getLogger(__name__)

    if request.method != "POST":
        logger.debug("add_account called with non-POST method: %s", request.method)
        return JsonResponse({"ok": False, "error": "POST required"}, status=405)
    username = (request.POST.get("username") or "").strip()
    password = (request.POST.get("password") or "").strip()
    # Diagnostic logging (do NOT log the raw password)
    try:
        logger.info(
            "add_account attempt user=%s session_key=%s",
            username,
            getattr(request.session, "session_key", None),
        )
        # Also log whether CSRF token was provided in POST and whether csrftoken cookie exists
        has_csrf_field = "csrfmiddlewaretoken" in request.POST
        csrf_cookie = request.COOKIES.get("csrftoken")
        logger.debug(
            "add_account csrf_field=%s csrf_cookie_present=%s alt_accounts_before=%s",
            has_csrf_field,
            bool(csrf_cookie),
            _get_alt_accounts(request.session),
        )
    except Exception:
        # best-effort logging; don't break the flow
        pass
    if not username or not password:
        logger.warning(
            "add_account missing username or password; username_present=%s",
            bool(username),
        )
        return JsonResponse(
            {"ok": False, "error": "username/password required"}, status=400
        )
    user = authenticate(request, username=username, password=password)
    if not user:
        logger.warning("add_account authentication failed for user=%s", username)
        return JsonResponse({"ok": False, "error": "invalid credentials"}, status=400)
    # store basic info in session (username and id)
    accounts = _get_alt_accounts(request.session)
    # avoid duplicates by username
    if any(a.get("username") == user.username for a in accounts):
        return JsonResponse(
            {"ok": True, "message": "already stored", "alt_accounts": accounts}
        )
    accounts.append({"username": user.username, "user_id": user.id})
    _set_alt_accounts(request.session, accounts)
    # return the updated alt_accounts for client-side verification
    return JsonResponse(
        {
            "ok": True,
            "username": user.username,
            "user_id": user.id,
            "alt_accounts": _get_alt_accounts(request.session),
        }
    )


@login_required
def remove_account(request):
    """Remove stored account from session. Expects POST with username."""
    if request.method != "POST":
        return JsonResponse({"ok": False, "error": "POST required"}, status=405)
    username = (request.POST.get("username") or "").strip()
    accounts = _get_alt_accounts(request.session)
    new = [a for a in accounts if a.get("username") != username]
    _set_alt_accounts(request.session, new)
    return JsonResponse({"ok": True})


@login_required
def switch_account(request):
    """Switch to a stored account by username. Uses login() to set request.user.
    Expects POST with username.
    """
    if request.method != "POST":
        return JsonResponse({"ok": False, "error": "POST required"}, status=405)
    username = (request.POST.get("username") or "").strip()
    accounts = _get_alt_accounts(request.session)
    target = next((a for a in accounts if a.get("username") == username), None)
    if not target:
        return JsonResponse({"ok": False, "error": "account not found"}, status=404)
    # Prepare to preserve current logged-in account so user can switch back.
    # We DO NOT write into the session before calling login() because Django
    # may rotate or replace the session during login(), which can discard
    # transient keys. Instead, build the accounts list here and write it
    # into the session after login() completes.
    try:
        cur = request.user
    except Exception:
        cur = None
    # copy existing accounts and ensure current user is included (if different)
    new_accounts = list(accounts) if accounts else []
    try:
        if cur and getattr(cur, "is_authenticated", False) and cur.username != username:
            if not any(a.get("username") == cur.username for a in new_accounts):
                new_accounts.append({"username": cur.username, "user_id": cur.id})
    except Exception:
        # fail-safe: ignore session changes if any error occurs
        pass
    # fetch user object
    from django.contrib.auth import get_user_model

    User = get_user_model()
    try:
        user = User.objects.get(pk=target.get("user_id"))
    except Exception:
        return JsonResponse({"ok": False, "error": "user missing"}, status=404)
    # perform login (this may rotate/replace the session)
    login(request, user)

    # persist the prepared alt-account list into the (post-login) session
    try:
        _set_alt_accounts(request.session, new_accounts)
    except Exception:
        # best-effort only
        pass

    # include alt_accounts after switch for diagnostics
    try:
        current_alt = _get_alt_accounts(request.session)
    except Exception:
        current_alt = []
    return JsonResponse(
        {"ok": True, "username": user.username, "alt_accounts": current_alt}
    )


# Small JSON API used by the frontend for realtime uniqueness checks
def api_check_equipment_id(request):
    """Return JSON { exists: bool, valid: bool, message: str }
    Query params: equipment_id, exclude (optional, id to exclude for edit forms)
    """
    # support both GET and POST (frontend may POST JSON)
    equipment_id = ""
    exclude = None
    if request.method == "POST":
        try:
            # If JSON body
            import json

            body = request.body.decode("utf-8") if request.body else ""
            if body:
                data = json.loads(body)
                equipment_id = (data.get("equipment_id") or "").strip()
                exclude = data.get("exclude")
        except Exception:
            # fallback to form-encoded
            equipment_id = (request.POST.get("equipment_id") or "").strip()
            exclude = request.POST.get("exclude")
    else:
        equipment_id = (request.GET.get("equipment_id") or "").strip()
        exclude = request.GET.get("exclude")

    result = {"exists": False, "valid": False, "message": ""}
    if equipment_id == "":
        result["valid"] = False
        result["message"] = "กรุณากรอกรหัส"
        return JsonResponse(result)

    if len(equipment_id) < 13:
        result["valid"] = False
        result["message"] = "รหัสต้องมีอย่างน้อย 13 ตัวอักษร"
        return JsonResponse(result)

    qs = Equipment_list.objects.filter(equipment_id=equipment_id)
    try:
        if exclude:
            qs = qs.exclude(id=int(exclude))
    except Exception:
        pass

    exists = qs.exists()
    result["exists"] = exists
    result["valid"] = not exists
    if exists:
        result["message"] = "รหัสนี้มีอยู่แล้ว"

    return JsonResponse(result)


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_schedule(request):
    """Schedule planned maintenance for equipment nearing PM due date.

    GET: show list of equipment with equipment_pm_due within `window_days` (default 30).
    POST: accept appointments: fields named schedule_<equipment_pk> = YYYY-MM-DD to create appointments.
    Capacity is enforced using MaintenanceCapacity model per date and equipment_type.
    """
    from datetime import timedelta

    from django.utils import timezone

    from .models import MaintenanceAppointment, MaintenanceCapacity

    window_days = int(request.GET.get("window", 30))
    today = timezone.localtime().date()
    end = today + timedelta(days=window_days)

    due_qs = Equipment_list.objects.filter(
        requires_pm=True, equipment_pm_due__gte=today, equipment_pm_due__lte=end
    ).order_by("equipment_pm_due")

    capacities = MaintenanceCapacity.objects.filter(date__gte=today, date__lte=end)
    cap_map = {}
    for c in capacities:
        cap_map.setdefault(c.date.isoformat(), {})[c.equipment_type] = c.capacity

    appts = MaintenanceAppointment.objects.filter(
        scheduled_date__gte=today, scheduled_date__lte=end
    )
    appt_count = {}
    for a in appts:
        key = f"{a.scheduled_date.isoformat()}||{a.equipment.equipment_type}"
        appt_count.setdefault(key, 0)
        appt_count[key] += 1

    if request.method == "POST":
        created = 0
        errors = []
        for eq in due_qs:
            key = f"schedule_{eq.pk}"
            val = request.POST.get(key)
            if not val:
                continue
            try:
                sel_date = timezone.datetime.strptime(val, "%Y-%m-%d").date()
                sel_key = sel_date.isoformat()
            except Exception:
                errors.append(f"วันที่ไม่ถูกต้องสำหรับ {eq.equipment_id}: {val}")
                continue

            # configured capacity
            cap = cap_map.get(sel_key, {}).get(eq.equipment_type, None)
            current = appt_count.get(f"{sel_key}||{eq.equipment_type}", 0)

            # compute available technicians for this equipment_type on sel_date
            try:
                from .models import Technician, TechnicianAvailability

                tech_qs = Technician.objects.filter(active=True)
                if eq.equipment_type:
                    # match technicians by normalized MasterItem skills
                    tech_qs = tech_qs.filter(
                        skills_m__label__icontains=eq.equipment_type
                    )
                tech_list = list(tech_qs)
                available_count = 0
                for t in tech_list:
                    av = TechnicianAvailability.objects.filter(
                        technician=t, date=sel_date
                    ).first()
                    if not av or av.status in ("working", "overtime"):
                        available_count += 1
            except Exception:
                available_count = 0

            per_tech_capacity = 1
            if cap is not None:
                estimated_capacity = (
                    min(cap, available_count * per_tech_capacity)
                    if available_count > 0
                    else 0
                )
            else:
                estimated_capacity = (
                    available_count * per_tech_capacity if available_count > 0 else 0
                )

            if estimated_capacity is not None and current >= estimated_capacity:
                errors.append(
                    f"ไม่สามารถนัด {eq.equipment_id} ใน {sel_date} ได้: เต็มตามความจุ/ช่างว่างสำหรับประเภท {eq.equipment_type}"
                )
                continue

            MaintenanceAppointment.objects.create(
                equipment=eq, scheduled_date=sel_date, created_by=request.user
            )
            appt_count[f"{sel_key}||{eq.equipment_type}"] = current + 1
            created += 1

        if created:
            messages.success(request, f"สร้างนัดหมายสำเร็จ {created} รายการ")
        if errors:
            for e in errors[:10]:
                messages.error(request, e)
        return redirect("maintenance_schedule")

    return render(
        request,
        "equipment/maintenance_schedule.html",
        {
            "due_list": due_qs,
            "due_rows": [
                {
                    "equipment": eq,
                    "capacity_samples": [
                        (
                            d,
                            cap_map.get(d, {}).get(eq.equipment_type, None),
                            appt_count.get(f"{d}||{eq.equipment_type}", 0),
                        )
                        for d in sorted(list(cap_map.keys()))
                    ],
                }
                for eq in due_qs
            ],
            "today": today,
            "end": end,
        },
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_capacity_list(request):
    """List and simple filter capacities. Provides link to add/create."""
    from .models import MaintenanceCapacity

    q = request.GET.get("q", "").strip()
    qs = MaintenanceCapacity.objects.all().order_by("date")
    if q:
        qs = qs.filter(equipment_type__icontains=q)
    return render(
        request, "equipment/maintenance_capacity_list.html", {"capacities": qs, "q": q}
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_capacity_create(request):
    from .models import MaintenanceCapacity

    if request.method == "POST":
        date = request.POST.get("date")
        equipment_type = request.POST.get("equipment_type", "").strip()
        cap = request.POST.get("capacity", "0")
        try:
            cap_i = int(cap)
        except Exception:
            cap_i = 0
        try:
            _ = MaintenanceCapacity.objects.create(
                date=date, equipment_type=equipment_type, capacity=cap_i
            )
            messages.success(request, "บันทึกความจุเรียบร้อย")
            return redirect("maintenance_capacity_list")
        except Exception as e:
            messages.error(request, f"Error: {e}")
            return render(
                request,
                "equipment/maintenance_capacity_form.html",
                {"form_values": request.POST},
            )
    return render(request, "equipment/maintenance_capacity_form.html", {})


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_capacity_edit(request, cap_id):
    from .models import MaintenanceCapacity

    try:
        c = MaintenanceCapacity.objects.get(id=cap_id)
    except MaintenanceCapacity.DoesNotExist:
        messages.error(request, "ไม่พบรายการ")
        return redirect("maintenance_capacity_list")
    if request.method == "POST":
        date = request.POST.get("date")
        equipment_type = request.POST.get("equipment_type", "").strip()
        cap = request.POST.get("capacity", "0")
        try:
            cap_i = int(cap)
        except Exception:
            cap_i = 0
        c.date = date
        c.equipment_type = equipment_type
        c.capacity = cap_i
        c.save()
        messages.success(request, "แก้ไขความจุเรียบร้อย")
        return redirect("maintenance_capacity_list")
    # GET
    form_values = {
        "date": c.date.isoformat(),
        "equipment_type": c.equipment_type,
        "capacity": c.capacity,
    }
    return render(
        request,
        "equipment/maintenance_capacity_form.html",
        {"form_values": form_values, "editing": True, "cap_id": c.id},
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_appointment_edit(request, appt_id):
    from .models import MaintenanceAppointment, MaintenanceCapacity

    try:
        a = MaintenanceAppointment.objects.select_related("equipment").get(id=appt_id)
    except MaintenanceAppointment.DoesNotExist:
        messages.error(request, "ไม่พบรายการ")
        return redirect("maintenance_appointments_list")

    if request.method == "POST":
        new_date = request.POST.get("scheduled_date")
        # capacity check
        eq_type = a.equipment.equipment_type
        cap_obj = MaintenanceCapacity.objects.filter(
            date=new_date, equipment_type=eq_type
        ).first()
        used = (
            MaintenanceAppointment.objects.filter(
                scheduled_date=new_date, equipment__equipment_type=eq_type
            )
            .exclude(pk=a.pk)
            .count()
        )
        if cap_obj and used >= cap_obj.capacity:
            messages.error(
                request, f"ไม่สามารถย้ายไป {new_date} ได้ — เต็มความจุสำหรับประเภท {eq_type}"
            )
            return render(
                request,
                "equipment/maintenance_appointment_edit.html",
                {"appt": a, "form_values": {"scheduled_date": new_date}},
            )
        a.scheduled_date = new_date
        a.save()
        messages.success(request, "บันทึกการเปลี่ยนแปลงนัดหมายแล้ว")
        return redirect("maintenance_appointments_list")

    # GET
    form_values = {"scheduled_date": a.scheduled_date.isoformat()}
    return render(
        request,
        "equipment/maintenance_appointment_edit.html",
        {"appt": a, "form_values": form_values},
    )


def api_check_capacity(request):
    """AJAX: ?date=YYYY-MM-DD&equipment_type=... -> {ok, capacity, used, remaining}

    remaining == null indicates unlimited (no capacity configured).
    """
    from .models import MaintenanceAppointment, MaintenanceCapacity

    date = request.GET.get("date")
    equipment_type = (request.GET.get("equipment_type") or "").strip()
    if not date or not equipment_type:
        return JsonResponse({"ok": False, "error": "missing parameters"}, status=400)
    cap_obj = MaintenanceCapacity.objects.filter(
        date=date, equipment_type=equipment_type
    ).first()
    used = MaintenanceAppointment.objects.filter(
        scheduled_date=date, equipment__equipment_type=equipment_type
    ).count()

    # Technician availability: count technicians whose skills match equipment_type (simple substring match)
    try:
        from .models import Technician, TechnicianAvailability

        tech_qs = Technician.objects.filter(active=True)
        if equipment_type:
            # prefer normalized skills_m lookup
            tech_qs = tech_qs.filter(skills_m__label__icontains=equipment_type)
        tech_list = list(tech_qs)
        # count those who are marked as working or overtime on that date; if no explicit availability record, assume working
        available_count = 0
        for t in tech_list:
            av = TechnicianAvailability.objects.filter(technician=t, date=date).first()
            if not av:
                # assume working by default
                available_count += 1
            else:
                if av.status in ("working", "overtime"):
                    available_count += 1
    except Exception:
        available_count = 0

    # per-technician capacity (devices per tech per day). Keep default 1 for now.
    per_tech_capacity = 1

    if cap_obj:
        configured = cap_obj.capacity
        # estimated capacity limited by both configured capacity and available technicians
        estimated_capacity = (
            min(configured, available_count * per_tech_capacity)
            if available_count > 0
            else 0
        )
    else:
        configured = None
        estimated_capacity = (
            available_count * per_tech_capacity if available_count > 0 else 0
        )

    remaining = None
    if estimated_capacity is not None:
        remaining = max(estimated_capacity - used, 0)

    return JsonResponse(
        {
            "ok": True,
            "capacity_configured": configured,
            "used": used,
            "available_techs": available_count,
            "per_tech_capacity": per_tech_capacity,
            "estimated_capacity": estimated_capacity,
            "remaining": remaining,
        }
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def technicians_list(request):
    from .models import Technician

    q = request.GET.get("q", "").strip()
    qs = Technician.objects.all().order_by("name")
    if q:
        qs = qs.filter(name__icontains=q)
    return render(
        request, "maintenance/technicians_list.html", {"technicians": qs, "q": q}
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def technician_create(request):
    from .forms import TechnicianForm

    if request.method == "POST":
        form = TechnicianForm(request.POST)
        if form.is_valid():
            tech = form.save(commit=False)
            tech.save()
            form.save_m2m()
            messages.success(request, "เพิ่มช่างเรียบร้อย")
            return redirect("technicians_list")
    else:
        form = TechnicianForm()
    return render(
        request,
        "maintenance/technician_form.html",
        {"form": form, "prefill_skills": []},
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def technician_edit(request, tech_id):
    from .forms import TechnicianForm
    from .models import Technician

    tech = Technician.objects.get(pk=tech_id)
    # Always prepare a prefill variable so it can be used when re-rendering the form
    prefill = []
    from .models import MasterItem

    if request.method == "POST":
        form = TechnicianForm(request.POST, instance=tech)
        if form.is_valid():
            tech = form.save(commit=False)
            tech.save()
            form.save_m2m()
            messages.success(request, "บันทึกข้อมูลช่างแล้ว")
            return redirect("technicians_list")
        else:
            # if the POST was invalid, try to preserve the submitted skills values
            try:
                posted = request.POST.getlist("skills_m")
                if posted:
                    items = MasterItem.objects.filter(id__in=posted)
                    prefill = [{"id": s.id, "text": s.label} for s in items]
            except Exception:
                prefill = []
    else:
        form = TechnicianForm(instance=tech)
        # prepare prefill data for Select2: list of {id, text}
        prefill = [{"id": s.id, "text": s.label} for s in tech.skills_m.all()]

    return render(
        request,
        "maintenance/technician_form.html",
        {"form": form, "edit": True, "technician": tech, "prefill_skills": prefill},
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def technician_delete(request, tech_id):
    from .models import Technician

    try:
        t = Technician.objects.get(pk=tech_id)
        t.delete()
        messages.success(request, "ลบช่างแล้ว")
    except Exception:
        messages.error(request, "ไม่พบช่าง")
    return redirect("technicians_list")


@permission_required("cmms.change_equipment_list", raise_exception=True)
def technician_availability_manage(request):
    """Simple manage page: list existing availability for a date and allow adding a record.

    POST params: technician_id, date, status, note
    """
    from datetime import date as date_cls

    from django.utils import timezone

    from .models import MasterItem, Technician, TechnicianAvailability

    # Support both date parameter and month/year parameters
    date_str = request.GET.get("date")
    month_str = request.GET.get("month")
    year_str = request.GET.get("year")

    if month_str and year_str:
        try:
            month = int(month_str)
            year = int(year_str)
            date = date_cls(year, month, 1)
        except Exception:
            date = timezone.localtime().date()
    elif date_str:
        try:
            date = timezone.datetime.strptime(date_str, "%Y-%m-%d").date()
        except Exception:
            date = timezone.localtime().date()
    else:
        date = timezone.localtime().date()

    # Load work statuses from master data
    work_statuses_qs = MasterItem.objects.filter(
        category="work_status", active=True
    ).order_by("order", "label")

    if request.method == "POST":
        # Support two POST formats:
        # 1) legacy: technician_id, status, note -> update single date
        # 2) calendar: fields named selected_day_<techid> = YYYY-MM-DD for each technician
        try:
            handled = False
            # Build a map of (tech_id, day) -> {status, overtime}
            calendar_data = {}

            # Process status radio buttons (working/off/holiday)
            for key, val in request.POST.items():
                if key.startswith("status_") and val:
                    parts = key.split("_")
                    if len(parts) < 3:
                        continue
                    try:
                        tid = int(parts[1])
                        day = int(parts[2])
                    except Exception:
                        continue
                    if (tid, day) not in calendar_data:
                        calendar_data[(tid, day)] = {"status": val, "overtime": False}
                    else:
                        calendar_data[(tid, day)]["status"] = val

            # Process overtime checkboxes
            for key, val in request.POST.items():
                if key.startswith("overtime_") and val:
                    parts = key.split("_")
                    if len(parts) < 3:
                        continue
                    try:
                        tid = int(parts[1])
                        day = int(parts[2])
                    except Exception:
                        continue
                    if (tid, day) not in calendar_data:
                        calendar_data[(tid, day)] = {
                            "status": "working",
                            "overtime": True,
                        }
                    else:
                        calendar_data[(tid, day)]["overtime"] = True

            # Process OT hidden inputs (template uses names like ot_<techId>_<day>)
            for key, val in request.POST.items():
                if key.startswith("ot_"):
                    parts = key.split("_")
                    if len(parts) < 3:
                        continue
                    try:
                        tid = int(parts[1])
                        day = int(parts[2])
                    except Exception:
                        continue
                    ot_code = val.strip() if isinstance(val, str) else val
                    if ot_code:
                        if (tid, day) not in calendar_data:
                            calendar_data[(tid, day)] = {
                                "status": "working",
                                "overtime": True,
                                "ot_code": ot_code,
                            }
                        else:
                            calendar_data[(tid, day)]["overtime"] = True
                            calendar_data[(tid, day)]["ot_code"] = ot_code

            # Debug: log calendar_data summary before saving
            try:
                logger.debug("Calendar POST data items=%d", len(calendar_data))
                # show a small sample to help debugging (max 10)
                sample = list(calendar_data.items())[:10]
                logger.debug("Calendar POST sample=%s", sample)
            except Exception:
                logger.exception("Failed to log calendar_data")

            # Save all calendar data
            try:
                cal_year = int(request.POST.get("calendar_year") or date.year)
                cal_month = int(request.POST.get("calendar_month") or date.month)
            except Exception:
                cal_year = date.year
                cal_month = date.month

            saved_count = 0
            for (tid, day), data in calendar_data.items():
                try:
                    wdate = date_cls(cal_year, cal_month, day)
                except Exception:
                    continue
                try:
                    tech = Technician.objects.get(pk=tid)
                except Technician.DoesNotExist:
                    continue

                # Determine final status: keep the base status (don't set 'overtime' as the status)
                # We persist OT separately (note or ot_code) so status values remain in the domain of work/leave
                final_status = data.get("status") or "working"

                # If OT code was provided (e.g., 'ot_4h'/'ot_7h'), store it in the note as legacy format
                note_val = ""
                if data and isinstance(data, dict) and data.get("ot_code"):
                    try:
                        note_val = f"OT:{data.get('ot_code')}"
                    except Exception:
                        note_val = str(data.get("ot_code"))

                TechnicianAvailability.objects.update_or_create(
                    technician=tech,
                    date=wdate,
                    defaults={"status": final_status, "note": note_val},
                )
                handled = True
                saved_count += 1

            if handled:
                messages.success(request, "บันทึกตารางการทำงานเรียบร้อย")
                logger.debug(
                    "Saved %d availability records from calendar POST", saved_count
                )
                # If caller requested to view report after save, redirect to report page
                if request.POST.get("redirect_to_report"):
                    try:
                        report_url = reverse("technician_availability_report")
                        return redirect(
                            f"{report_url}?month={cal_month}&year={cal_year}"
                        )
                    except Exception:
                        return redirect("technician_availability_report")
                return redirect("technician_availability_manage")

            # legacy single-update fallback
            tech_id = request.POST.get("technician_id")
            if tech_id:
                status = request.POST.get("status", "working")
                note = request.POST.get("note", "")
                tech = Technician.objects.get(pk=int(tech_id))
                obj, created = TechnicianAvailability.objects.update_or_create(
                    technician=tech,
                    date=date,
                    defaults={"status": status, "note": note},
                )
                messages.success(request, "บันทึกสถานะช่างเรียบร้อย")
                return redirect("technician_availability_manage")
        except Exception as e:
            messages.error(request, f"Error: {e}")

    # List availabilities for the entire month
    avails_qs = TechnicianAvailability.objects.filter(
        date__year=date.year, date__month=date.month
    ).select_related("technician")
    technicians_qs = Technician.objects.filter(active=True).order_by("name")

    # Debug: Check if we have technicians — use logging instead of print to control verbosity
    try:
        tech_count = technicians_qs.count()
    except Exception:
        tech_count = len(list(technicians_qs))
    # log count and a small sample (max 20) to avoid huge dumps in terminal
    tech_sample = [{"id": t.id, "name": t.name} for t in technicians_qs[:20]]
    logger.debug("Found %d active technicians; sample=%s", tech_count, tech_sample)

    # Serialize into JSON-safe Python structures for template json_script
    technicians_serial = [
        {"id": t.id, "name": t.name, "phone": t.phone} for t in technicians_qs
    ]
    logger.debug("Serialized technicians count=%d", len(technicians_serial))

    availabilities_serial = []
    for a in avails_qs:
        try:
            availabilities_serial.append(
                {
                    "id": a.id,
                    "technician": {"id": a.technician.id, "name": a.technician.name},
                    "work_date": a.date.isoformat(),
                    "status": a.status,
                    "note": a.note,
                }
            )
        except Exception:
            continue

    # Serialize work statuses from master data with parent-child structure
    work_statuses_serial = []
    for ws in work_statuses_qs:
        label_lower = (ws.label or "").lower()

        # Get children (sub-types)
        children = []
        for child in ws.children.filter(active=True).order_by("order", "label"):
            children.append(
                {
                    "id": child.id,
                    "name": child.label,
                    "code": child.code or child.label.lower().replace(" ", "_"),
                    "color": child.description or ws.description or "secondary",
                    "is_default": bool(getattr(child, "is_default", False)),
                }
            )

        work_statuses_serial.append(
            {
                "id": ws.id,
                "name": ws.label,
                "code": ws.code or label_lower.replace(" ", "_"),
                "color": ws.description or "secondary",
                "is_overtime": "overtime" in label_lower
                or "ot" in label_lower
                or "ล่วงเวลา" in ws.label,
                "is_leave": "leave" in label_lower or "ลา" in ws.label,
                "children": children,
            }
        )

    # Get shift schedule data for this month
    from .models import ShiftSchedule

    shift_schedules_qs = ShiftSchedule.objects.filter(
        date__year=date.year, date__month=date.month
    ).select_related("technician")

    # Debug: แสดงข้อมูล query
    print(f"🔍 Querying ShiftSchedule for year={date.year}, month={date.month}")
    print(f"📊 Found {shift_schedules_qs.count()} schedules")

    shift_schedules_serial = {}
    for schedule in shift_schedules_qs:
        date_key = schedule.date.isoformat()
        if date_key not in shift_schedules_serial:
            shift_schedules_serial[date_key] = {}
        shift_schedules_serial[date_key][schedule.shift] = {
            "technician_id": schedule.technician.id,
            "technician_name": schedule.technician.name,
            "note": schedule.note,
        }

    print(f"✅ Serialized shift_schedules keys: {list(shift_schedules_serial.keys())}")

    return render(
        request,
        "maintenance/technician_availability_manage.html",
        {
            "date": date,
            "availabilities": availabilities_serial,
            "technicians": technicians_serial,
            "work_statuses": work_statuses_serial,
            "shift_schedules": shift_schedules_serial,
        },
    )


def technician_availability_manage_v2(request):
    """Modern UI version of technician availability management."""
    from datetime import date as date_cls

    from django.utils import timezone

    from .models import MasterItem, Technician, TechnicianAvailability

    # Get month/year parameters
    month_str = request.GET.get("month")
    year_str = request.GET.get("year")

    if month_str and year_str:
        try:
            month = int(month_str)
            year = int(year_str)
        except Exception:
            today = timezone.localtime().date()
            month = today.month
            year = today.year
    else:
        today = timezone.localtime().date()
        month = today.month
        year = today.year

    if request.method == "POST":
        # Process form submission
        try:
            post_month = int(request.POST.get("month", month))
            post_year = int(request.POST.get("year", year))

            # Build calendar data from POST
            calendar_data = {}

            # Process status fields (work/leave)
            for key, val in request.POST.items():
                if key.startswith("status_") and val:
                    parts = key.split("_")
                    if len(parts) >= 3:
                        try:
                            tid = int(parts[1])
                            day = int(parts[2])
                            if (tid, day) not in calendar_data:
                                calendar_data[(tid, day)] = {}
                            calendar_data[(tid, day)]["status"] = val
                        except (ValueError, IndexError):
                            continue

            # Process OT fields
            for key, val in request.POST.items():
                if key.startswith("ot_") and val:
                    parts = key.split("_")
                    if len(parts) >= 3:
                        try:
                            tid = int(parts[1])
                            day = int(parts[2])
                            if (tid, day) not in calendar_data:
                                calendar_data[(tid, day)] = {}
                            calendar_data[(tid, day)]["ot"] = val
                        except (ValueError, IndexError):
                            continue

            # Save to database
            for (tid, day), data in calendar_data.items():
                try:
                    wdate = date_cls(post_year, post_month, day)
                    tech = Technician.objects.get(pk=tid)

                    # Prepare defaults
                    defaults = {"status": data.get("status", "work"), "note": ""}

                    # If there's an OT field, save it separately (may need to add ot_type field to model)
                    # For now, we can store it in the note or extend the model
                    if "ot" in data and data["ot"]:
                        defaults["note"] = f"OT: {data['ot']}"

                    TechnicianAvailability.objects.update_or_create(
                        technician=tech, date=wdate, defaults=defaults
                    )
                except Exception:
                    continue

            messages.success(request, "บันทึกตารางการทำงานเรียบร้อยแล้ว")
            return redirect(f"{request.path}?month={post_month}&year={post_year}")

        except Exception as e:
            messages.error(request, f"เกิดข้อผิดพลาด: {e}")

    # Load data
    technicians_qs = Technician.objects.filter(active=True).order_by(
        "first_name", "last_name"
    )

    avails_qs = TechnicianAvailability.objects.filter(
        date__year=year, date__month=month
    ).select_related("technician")

    work_statuses_qs = MasterItem.objects.filter(
        category="work_status", active=True
    ).order_by("order", "label")

    # Serialize technicians
    technicians_serial = []
    for t in technicians_qs:
        technicians_serial.append(
            {
                "id": t.id,
                "first_name": t.first_name,
                "last_name": t.last_name,
                "name": f"{t.first_name} {t.last_name}",
                "phone": t.phone,
            }
        )

    # Serialize availabilities
    availabilities_serial = []
    for a in avails_qs:
        try:
            availabilities_serial.append(
                {
                    "id": a.id,
                    "technician": {
                        "id": a.technician.id,
                        "name": f"{a.technician.first_name} {a.technician.last_name}",
                    },
                    "work_date": a.date.isoformat(),
                    "status": a.status,
                    "note": a.note,
                    "ot_type": None,  # TODO: Add ot_type field to model
                }
            )
        except Exception:
            continue

    # Serialize work statuses with parent-child
    work_statuses_serial = []
    for ws in work_statuses_qs:
        label_lower = (ws.label or "").lower()

        children = []
        for child in ws.children.filter(active=True).order_by("order", "label"):
            children.append(
                {
                    "id": child.id,
                    "name": child.label,
                    "code": child.code or child.label.lower().replace(" ", "_"),
                    "display_name": child.label,
                    "is_default": bool(getattr(child, "is_default", False)),
                }
            )

        work_statuses_serial.append(
            {
                "id": ws.id,
                "name": ws.label,
                "code": ws.code or label_lower.replace(" ", "_"),
                "display_name": ws.label,
                "is_overtime": "overtime" in label_lower
                or "ot" in label_lower
                or "ล่วงเวลา" in ws.label,
                "is_leave": "leave" in label_lower
                or "ลา" in ws.label
                or "หยุด" in ws.label,
                "children": children,
            }
        )

    return render(
        request,
        "maintenance/technician_availability_manage_v2.html",
        {
            "selected_month": month,
            "selected_year": year,
            "availabilities": availabilities_serial,
            "technicians": technicians_serial,
            "work_statuses": work_statuses_serial,
        },
    )


def technician_availability_report(request):
    """Report page: read-only view of technician availability with statistics."""
    from datetime import date as date_cls

    from django.utils import timezone

    from .models import Technician, TechnicianAvailability

    # Support date parameter for single-day filtering (used by daily_capacity.html)
    date_param = request.GET.get("date", None)
    month_param = request.GET.get("month", None)
    year_param = request.GET.get("year", None)

    # If date parameter is provided, use it for single-day filtering
    specific_date = None
    if date_param:
        try:
            from datetime import datetime
            specific_date = datetime.strptime(date_param, "%Y-%m-%d").date()
            year = specific_date.year
            month = specific_date.month
            date = specific_date
        except ValueError:
            specific_date = None

    # Otherwise use month/year parameters
    if not specific_date:
        try:
            year = (
                int(year_param)
                if (year_param is not None and year_param != "")
                else timezone.localtime().year
            )
        except ValueError:
            year = timezone.localtime().year

        # Interpret an omitted month as the current month. An explicit empty
        # string ('') means "all months" (returned as month=None below).
        if month_param is None:
            month = timezone.localtime().month
        elif month_param == "":
            month = None
        else:
            try:
                month = int(month_param)
            except ValueError:
                month = timezone.localtime().month

        # Create a date object for template header. If month is None use Jan 1
        # of the year (template can still render a year header while data is
        # for the whole year).
        if month is None:
            date = date_cls(year, 1, 1)
        else:
            date = date_cls(year, month, 1)

    # Get all technicians (used for the technician selector in the template)
    technicians_qs = Technician.objects.filter(active=True).order_by("name")

    # Optional: filter by a selected technician id passed via GET
    selected_technician = None
    tech_param = (request.GET.get("technician") or "").strip()
    try:
        if tech_param != "":
            selected_technician = int(tech_param)
    except Exception:
        selected_technician = None

    # Get availabilities - filter by specific date if provided, otherwise by month/year
    if specific_date:
        avails_qs = TechnicianAvailability.objects.filter(date=specific_date)
    else:
        avails_qs = TechnicianAvailability.objects.filter(date__year=year)
        if month is not None:
            avails_qs = avails_qs.filter(date__month=month)
    avails_qs = avails_qs.select_related("technician")
    if selected_technician:
        avails_qs = avails_qs.filter(technician_id=selected_technician)

    # Serialize data
    technicians_serial = []
    for t in technicians_qs:
        # Get skills from ManyToMany or text field
        skills_list = list(t.skills_m.values_list('label', flat=True)) if t.skills_m.exists() else []
        skills_str = ", ".join(skills_list) if skills_list else (t.skills or "")
        
        technicians_serial.append({
            "id": t.id,
            "name": t.name,
            "skills": skills_str,
            "capacity": t.per_day_capacity,
        })
    
    availabilities_serial = []
    for a in avails_qs:
        try:
            # Get skills for this technician
            skills_list = list(a.technician.skills_m.values_list('label', flat=True)) if a.technician.skills_m.exists() else []
            skills_str = ", ".join(skills_list) if skills_list else (a.technician.skills or "")
            
            availabilities_serial.append(
                {
                    "id": a.id,
                    "technician": {
                        "id": a.technician.id,
                        "name": a.technician.name,
                        "skills": skills_str,
                        "capacity": a.technician.per_day_capacity,
                    },
                    "work_date": a.date.isoformat(),
                    "status": a.status,
                    "note": a.note,
                }
            )
        except Exception:
            continue

    # Serialize shift schedules for the month so the report can display shifts and compute shift hours
    try:
        from .models import ShiftSchedule

        # Filter shift schedules - by specific date if provided, otherwise by month/year
        if specific_date:
            schedules_qs = ShiftSchedule.objects.filter(date=specific_date)
        else:
            schedules_qs = ShiftSchedule.objects.filter(date__year=year)
            if month is not None:
                schedules_qs = schedules_qs.filter(date__month=month)
        if selected_technician:
            schedules_qs = schedules_qs.filter(technician_id=selected_technician)
        schedules_qs = schedules_qs.select_related("technician")
        shift_schedules = {}
        for s in schedules_qs:
            date_key = s.date.strftime("%Y-%m-%d")
            if date_key not in shift_schedules:
                shift_schedules[date_key] = []
            shift_schedules[date_key].append(
                {
                    "id": s.id,
                    "technician_id": s.technician.id if s.technician else None,
                    "technician_name": s.technician.name if s.technician else None,
                    "shift": s.shift,
                    "note": s.note,
                }
            )
    except Exception:
        shift_schedules = {}

    return render(
        request,
        "maintenance/technician_availability_report.html",
        {
            "date": date,
            "availabilities": availabilities_serial,
            "technicians": technicians_serial,
            "shift_schedules": shift_schedules,
            "selected_technician": selected_technician,
            # pass selected_month so the template/JS can detect all-months (None)
            "selected_month": month,
        },
    )


def shift_schedule(request):
    """Shift schedule management page for assigning technicians to shifts."""
    import json
    from datetime import datetime

    from django.db.models import Case, IntegerField, When

    from .models import ShiftSchedule, Technician

    # Get year and month from request, default to current
    year = int(request.GET.get("year", datetime.now().year))
    month = int(request.GET.get("month", datetime.now().month))

    # Define custom order for technicians
    custom_order = ["ลิขิต", "เอกพันธ์", "อมรเทพ", "อทิตย์", "อนุรักษ์", "บุญฤทธิ์"]

    # Create CASE WHEN conditions for custom ordering
    ordering_conditions = [
        When(name__icontains=name, then=pos) for pos, name in enumerate(custom_order)
    ]

    # Get active technicians who can work shifts, ordered by custom order then by name
    technicians = (
        Technician.objects.filter(active=True, can_work_shifts=True)
        .annotate(
            custom_order=Case(
                *ordering_conditions,
                default=len(custom_order),
                output_field=IntegerField(),
            )
        )
        .order_by("custom_order", "name")
    )

    # Convert to JSON format
    technicians_list = [{"id": t.id, "name": t.name} for t in technicians]
    technicians_json = json.dumps(technicians_list, ensure_ascii=False)

    # Get existing schedules for this month
    existing_schedules = ShiftSchedule.objects.filter(
        date__year=year, date__month=month
    ).select_related("technician")

    # Convert to JSON format for JavaScript
    schedules_data = {}
    for schedule in existing_schedules:
        date_key = schedule.date.strftime("%Y-%m-%d")
        if date_key not in schedules_data:
            schedules_data[date_key] = {}
        schedules_data[date_key][schedule.shift] = {
            "technician_id": schedule.technician.id,
            "technician_name": schedule.technician.name,
            "note": schedule.note,
        }
    schedules_json = json.dumps(schedules_data, ensure_ascii=False)

    return render(
        request,
        "maintenance/shift_schedule.html",
        {
            "technicians": technicians,
            "technicians_json": technicians_json,
            "schedules_json": schedules_json,
            "current_year": year,
            "current_month": month,
        },
    )


@login_required
@require_http_methods(["POST"])
def api_save_shift_schedule(request):
    """API endpoint for saving shift schedule."""
    import json

    from django.db import transaction

    from .models import ShiftSchedule, Technician

    try:
        data = json.loads(request.body)
        schedules = data.get("schedules", [])

        if not schedules:
            return JsonResponse(
                {"success": False, "error": "ไม่มีข้อมูลที่จะบันทึก"}, status=400
            )

        saved_count = 0
        updated_count = 0
        errors = []

        with transaction.atomic():
            for item in schedules:
                try:
                    date = item.get("date")
                    shift = item.get("shift")
                    technician_id = item.get("technician_id")

                    if not all([date, shift, technician_id]):
                        errors.append(f"ข้อมูลไม่ครบ: {item}")
                        continue

                    # ตรวจสอบว่าช่างมีอยู่จริง
                    try:
                        technician = Technician.objects.get(
                            id=technician_id, active=True
                        )
                    except Technician.DoesNotExist:
                        errors.append(f"ไม่พบช่าง ID: {technician_id}")
                        continue

                    # บันทึกหรืออัพเดท
                    schedule, created = ShiftSchedule.objects.update_or_create(
                        date=date,
                        shift=shift,
                        defaults={
                            "technician": technician,
                            "created_by": request.user,
                        },
                    )

                    if created:
                        saved_count += 1
                    else:
                        updated_count += 1

                except Exception as e:
                    errors.append(f"Error processing {item}: {str(e)}")

        return JsonResponse(
            {
                "success": True,
                "message": f"บันทึกสำเร็จ {saved_count} รายการ, อัพเดท {updated_count} รายการ",
                "saved_count": saved_count,
                "updated_count": updated_count,
                "errors": errors if errors else None,
            }
        )

    except json.JSONDecodeError:
        return JsonResponse(
            {"success": False, "error": "รูปแบบ JSON ไม่ถูกต้อง"}, status=400
        )
    except Exception as e:
        return JsonResponse(
            {"success": False, "error": f"เกิดข้อผิดพลาด: {str(e)}"}, status=500
        )


@login_required
def shift_schedule_report(request):
    """View shift schedule report."""
    import calendar
    from datetime import datetime

    from .models import ShiftSchedule

    # Get month and year from query params (default to current month)
    today = datetime.now().date()
    year = int(request.GET.get("year", today.year))
    month = int(request.GET.get("month", today.month))

    # Calculate date range
    first_day = datetime(year, month, 1).date()
    last_day = datetime(year, month, calendar.monthrange(year, month)[1]).date()

    # Get all schedules for this month
    schedules = (
        ShiftSchedule.objects.filter(date__gte=first_day, date__lte=last_day)
        .select_related("technician", "created_by")
        .order_by("date", "shift")
    )

    # Group by date
    schedule_by_date = {}
    for schedule in schedules:
        date_str = schedule.date.strftime("%Y-%m-%d")
        if date_str not in schedule_by_date:
            schedule_by_date[date_str] = {}
        schedule_by_date[date_str][schedule.shift] = {
            "id": schedule.id,
            "technician": schedule.technician,
            "created_at": schedule.created_at,
            "created_by": schedule.created_by,
        }

    # Calculate prev/next month
    if month == 1:
        prev_month, prev_year = 12, year - 1
    else:
        prev_month, prev_year = month - 1, year

    if month == 12:
        next_month, next_year = 1, year + 1
    else:
        next_month, next_year = month + 1, year

    # Thai months array
    thai_months = [
        "มกราคม",
        "กุมภาพันธ์",
        "มีนาคม",
        "เมษายน",
        "พฤษภาคม",
        "มิถุนายน",
        "กรกฎาคม",
        "สิงหาคม",
        "กันยายน",
        "ตุลาคม",
        "พฤศจิกายน",
        "ธันวาคม",
    ]

    context = {
        "year": year,
        "month": month,
        "month_name": thai_months[month - 1],  # Get month name (1-based to 0-based)
        "first_day": first_day,
        "last_day": last_day,
        "schedule_by_date": schedule_by_date,
        "prev_month": prev_month,
        "prev_year": prev_year,
        "next_month": next_month,
        "next_year": next_year,
        "thai_months": thai_months,
    }

    return render(request, "maintenance/shift_schedule_report.html", context)


@login_required
@require_http_methods(["DELETE"])
def api_delete_shift_schedule(request, schedule_id):
    """API endpoint for deleting a shift schedule."""
    from .models import ShiftSchedule

    try:
        schedule = ShiftSchedule.objects.get(id=schedule_id)

        # Store info for response
        date_str = schedule.date.strftime("%Y-%m-%d")
        shift_name = "เวรเช้า" if schedule.shift == "morning" else "เวรบ่าย"

        # Delete
        schedule.delete()

        return JsonResponse(
            {"success": True, "message": f"ลบ {shift_name} วันที่ {date_str} เรียบร้อยแล้ว"}
        )

    except ShiftSchedule.DoesNotExist:
        return JsonResponse(
            {"success": False, "error": "ไม่พบข้อมูลที่ต้องการลบ"}, status=404
        )

    except Exception as e:
        return JsonResponse({"success": False, "error": str(e)}, status=500)


def api_technicians_by_date(request):
    """Public-ish API: ?date=YYYY-MM-DD&skill=optional -> list technicians and their status for that date."""
    from .models import Technician, TechnicianAvailability

    date = request.GET.get("date")
    skill = (request.GET.get("skill") or "").strip()
    if not date:
        return JsonResponse({"ok": False, "error": "missing date"}, status=400)
    qs = Technician.objects.filter(active=True)
    if skill:
        from django.db.models import Q

        # if skill is numeric id, filter by skills_m id; otherwise match either skills_m label or legacy skills text
        try:
            sid = int(skill)
            qs = qs.filter(Q(skills_m__id=sid) | Q(skills__icontains=skill))
        except Exception:
            qs = qs.filter(
                Q(skills_m__label__icontains=skill) | Q(skills__icontains=skill)
            )
    techs = []
    for t in qs.order_by("name"):
        av = TechnicianAvailability.objects.filter(technician=t, date=date).first()
        skills_labels = [mi.label for mi in t.skills_m.all()]
        techs.append(
            {
                "id": t.id,
                "name": t.name,
                "skills": skills_labels,
                "phone": t.phone,
                "status": av.status if av else "working",
                "note": av.note if av else "",
            }
        )
    return JsonResponse({"ok": True, "date": date, "technicians": techs})


def api_equipment_for_select2(request):
    """Return equipment options for select2: ?q=search"""
    q = (request.GET.get("q") or "").strip()
    from .models import Equipment_list

    qs = Equipment_list.objects.all()
    if q:
        qs = (
            qs.filter(equipment_name_TH__icontains=q)
            | qs.filter(equipment_name_EN__icontains=q)
            | qs.filter(equipment_code__icontains=q)
        )
    results = []
    for e in qs.order_by("equipment_name_TH")[:50]:
        # id and text expected by select2
        results.append(
            {
                "id": e.equipment_name_TH or e.equipment_name_EN or e.equipment_code,
                "text": f"{e.equipment_name_TH} ({e.equipment_code})",
            }
        )
    return JsonResponse({"results": results})


def api_master_skills_for_select2(request):
    """Return MasterItem options for select2 for category 'technician_skill'"""
    q = (request.GET.get("q") or "").strip()
    # allow optional category param; default to 'technician_skill' for backwards compatibility
    category = (request.GET.get("category") or "technician_skill").strip()
    from .models import MasterItem

    qs = MasterItem.objects.filter(category=category, active=True)
    if q:
        qs = qs.filter(label__icontains=q)
    results = [
        {"id": m.id, "text": m.label} for m in qs.order_by("order", "label")[:200]
    ]
    return JsonResponse({"results": results})


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_capacity_delete(request, cap_id):
    from .models import MaintenanceCapacity

    try:
        c = MaintenanceCapacity.objects.get(id=cap_id)
        c.delete()
        messages.success(request, "ลบความจุเรียบร้อย")
    except Exception:
        messages.error(request, "ไม่พบรายการ")
    return redirect("maintenance_capacity_list")


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_appointments_list(request):
    from .models import (
        MaintenanceAppointment,
        Equipment_list,
        ServiceRequest,
    )
    from django.utils import timezone
    from django.db.models import Q, OuterRef, Exists

    today = timezone.localtime().date()

    # Base querysets
    appt_qs = MaintenanceAppointment.objects.select_related(
        "equipment", "created_by"
    )

    # ServiceRequests that specify a preferred_service_date and have equipment
    # Exclude deleted SRs and SRs that already have a matching MaintenanceAppointment
    sr_qs = (
        ServiceRequest.objects.select_related("equipment", "requested_by")
        .filter(preferred_service_date__isnull=False, equipment__isnull=False)
        .exclude(status="deleted")
    )

    # Annotate SRs that already have an appointment on the same equipment+date
    existing_appt = MaintenanceAppointment.objects.filter(
        equipment=OuterRef("equipment"), scheduled_date=OuterRef("preferred_service_date")
    )
    sr_qs = sr_qs.annotate(has_appt=Exists(existing_appt)).filter(has_appt=False)

    # --- filters from request ---
    q = request.GET.get("q", "").strip()
    date_filter = request.GET.get("date_filter", "upcoming")  # upcoming | today | past | all
    eq_type_filter = request.GET.get("eq_type", "").strip()

    if q:
        appt_qs = appt_qs.filter(
            Q(equipment__equipment_id__icontains=q)
            | Q(equipment__equipment_name_TH__icontains=q)
            | Q(equipment__equipment_name_EN__icontains=q)
            | Q(equipment__equipment_user_customer__icontains=q)
        )
        sr_qs = sr_qs.filter(
            Q(equipment__equipment_id__icontains=q)
            | Q(equipment__equipment_name_TH__icontains=q)
            | Q(equipment__equipment_name_EN__icontains=q)
            | Q(customer_name__icontains=q)
        )

    if eq_type_filter:
        appt_qs = appt_qs.filter(equipment__equipment_type__icontains=eq_type_filter)
        sr_qs = sr_qs.filter(equipment__equipment_type__icontains=eq_type_filter)

    if date_filter == "today":
        appt_qs = appt_qs.filter(scheduled_date=today)
        sr_qs = sr_qs.filter(preferred_service_date=today)
    elif date_filter == "upcoming":
        appt_qs = appt_qs.filter(scheduled_date__gte=today)
        sr_qs = sr_qs.filter(preferred_service_date__gte=today)
    elif date_filter == "past":
        appt_qs = appt_qs.filter(scheduled_date__lt=today)
        sr_qs = sr_qs.filter(preferred_service_date__lt=today)

    # Build combined list of dicts for template (so SRs and Appointments render uniformly)
    results = []
    for a in appt_qs.order_by("scheduled_date"):
        results.append(
            {
                "source": "appt",
                "id": a.id,
                "scheduled_date": a.scheduled_date,
                "equipment": a.equipment,
                "notes": a.notes,
                "created_by": a.created_by,
                "obj": a,
            }
        )

    for sr in sr_qs.order_by("preferred_service_date"):
        results.append(
            {
                "source": "sr",
                "id": sr.id,
                "scheduled_date": sr.preferred_service_date,
                "equipment": sr.equipment,
                "notes": sr.notes,
                "created_by": sr.requested_by,
                "obj": sr,
            }
        )

    # Sort combined results by date then equipment id
    from datetime import date as _date

    def _sort_key(item):
        sd = item.get("scheduled_date") or _date.max
        eqid = "" if not item.get("equipment") else (item["equipment"].equipment_id or "")
        return (sd, eqid)

    results.sort(key=_sort_key)

    # Compute stats that include SRs with preferred dates (not yet scheduled)
    total_appts_db = MaintenanceAppointment.objects.count()
    total_srs_with_pref = (
        ServiceRequest.objects.filter(preferred_service_date__isnull=False, equipment__isnull=False)
        .exclude(status="deleted")
        .annotate(has_appt=Exists(existing_appt))
        .filter(has_appt=False)
        .count()
    )

    stats = {
        "total": total_appts_db + total_srs_with_pref,
        "today": (
            MaintenanceAppointment.objects.filter(scheduled_date=today).count()
            + ServiceRequest.objects.filter(preferred_service_date=today).exclude(status="deleted").count()
        ),
        "upcoming": (
            MaintenanceAppointment.objects.filter(scheduled_date__gt=today).count()
            + ServiceRequest.objects.filter(preferred_service_date__gt=today).exclude(status="deleted").count()
        ),
        "past": (
            MaintenanceAppointment.objects.filter(scheduled_date__lt=today).count()
            + ServiceRequest.objects.filter(preferred_service_date__lt=today).exclude(status="deleted").count()
        ),
    }

    # Equipment types dropdown (from both appointments and SRs)
    eq_types_qs = (
        Equipment_list.objects.filter(
            Q(maintenance_appointments__isnull=False) | Q(servicerequest__isnull=False)
        )
        .values_list("equipment_type", flat=True)
        .distinct()
        .order_by("equipment_type")
    )
    eq_types = [t for t in eq_types_qs if t and t.strip()]

    return render(
        request,
        "equipment/maintenance_appointments_list.html",
        {
            "appointments": results,
            "today": today,
            "q": q,
            "date_filter": date_filter,
            "eq_type_filter": eq_type_filter,
            "eq_types": eq_types,
            "stats": stats,
        },
    )


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_appointment_delete(request, appt_id):
    from .models import MaintenanceAppointment

    try:
        a = MaintenanceAppointment.objects.get(id=appt_id)
        a.delete()
        messages.success(request, "ยกเลิกนัดหมายเรียบร้อย")
    except Exception:
        messages.error(request, "ไม่พบรายการ")
    return redirect("maintenance_appointments_list")


@permission_required("cmms.change_equipment_list", raise_exception=True)
def maintenance_technician_schedule(request):
    """Show appointment schedule grouped by technician."""
    from .models import MaintenanceAppointment, Technician
    from django.utils import timezone

    today = timezone.localtime().date()
    selected_tech_id = request.GET.get("tech_id", "").strip()
    date_filter = request.GET.get("date_filter", "upcoming")

    # ── per-technician appointment summary ───────────────────────────────────
    active_techs = Technician.objects.filter(active=True).order_by("name")
    tech_summaries = []
    for tech in active_techs:
        base = MaintenanceAppointment.objects.filter(assigned_technician=tech)
        tech_summaries.append(
            {
                "tech": tech,
                "total": base.count(),
                "today": base.filter(scheduled_date=today).count(),
                "upcoming": base.filter(scheduled_date__gt=today).count(),
                "past": base.filter(scheduled_date__lt=today).count(),
            }
        )

    # ── unassigned appointments ───────────────────────────────────────────────
    unassigned_count = MaintenanceAppointment.objects.filter(
        assigned_technician__isnull=True
    ).count()

    # ── appointments for the selected technician ─────────────────────────────
    selected_tech = None
    appointments = []
    if selected_tech_id == "unassigned":
        appt_qs = MaintenanceAppointment.objects.filter(
            assigned_technician__isnull=True
        ).select_related("equipment", "created_by")
    elif selected_tech_id:
        try:
            selected_tech = Technician.objects.get(id=int(selected_tech_id))
        except (Technician.DoesNotExist, ValueError):
            selected_tech = None
        if selected_tech:
            appt_qs = MaintenanceAppointment.objects.filter(
                assigned_technician=selected_tech
            ).select_related("equipment", "created_by")
        else:
            appt_qs = MaintenanceAppointment.objects.none()
    else:
        appt_qs = MaintenanceAppointment.objects.none()

    if selected_tech_id:
        if date_filter == "today":
            appt_qs = appt_qs.filter(scheduled_date=today)
        elif date_filter == "upcoming":
            appt_qs = appt_qs.filter(scheduled_date__gte=today)
        elif date_filter == "past":
            appt_qs = appt_qs.filter(scheduled_date__lt=today)
        appointments = list(appt_qs.order_by("scheduled_date"))

    return render(
        request,
        "equipment/maintenance_technician_schedule.html",
        {
            "tech_summaries": tech_summaries,
            "selected_tech": selected_tech,
            "selected_tech_id": selected_tech_id,
            "appointments": appointments,
            "today": today,
            "date_filter": date_filter,
            "unassigned_count": unassigned_count,
        },
    )

@login_required
def equipment_maintenance_schedule(request):
    """Display equipment maintenance schedule page with due equipment list."""

    from django.utils import timezone

    today = timezone.localtime().date()
    current_year = today.year
    current_month = today.month

    # Get unique departments/customers from equipment list
    departments = (
        Equipment_list.objects.filter(requires_pm=True)
        .values_list("equipment_user_customer", flat=True)
        .distinct()
        .order_by("equipment_user_customer")
    )

    # Filter out empty departments
    departments = [d for d in departments if d and d.strip()]

    context = {
        "current_year": current_year,
        "current_month": current_month,
        "today": today,
        "departments": departments,
    }
    return render(request, "maintenance/equipment_maintenance_schedule.html", context)


@login_required
def api_equipment_due_list(request):
    """API endpoint to get equipment list that are due for maintenance.

    Query parameters:
    - month: Filter by month (1-12)
    - year: Filter by year
    - department: Filter by department
    - status: Filter by status (overdue, due-soon, scheduled)
    - search: Search by equipment_id or equipment_name_TH

    Returns JSON with equipment list and their maintenance status.
    """
    from datetime import datetime

    from django.db.models import Q
    from django.utils import timezone

    from .models import MaintenanceAppointment

    try:
        today = timezone.localtime().date()

        # Get filter parameters
        month = request.GET.get("month", "")
        year = request.GET.get("year", str(today.year))
        department = request.GET.get("department", "")
        status_filter = request.GET.get("status", "")
        search = request.GET.get("search", "").strip()

        # Determine whether caller wants PM+CAL combined results
        pm_cal_flag = request.GET.get('pm_cal_due', '')

        # Base queryset - include PM and/or CAL depending on flag
        if pm_cal_flag:
            from django.db.models import Q

            qs = Equipment_list.objects.filter(
                Q(requires_pm=True, equipment_pm_due__isnull=False)
                | Q(requires_cal=True, equipment_cal_due__isnull=False)
            )
        else:
            qs = Equipment_list.objects.filter(requires_pm=True)

        # Apply filters
        if search:
            qs = qs.filter(
                Q(equipment_id__icontains=search)
                | Q(equipment_name_TH__icontains=search)
                | Q(equipment_name_EN__icontains=search)
            )

        if department:
            qs = qs.filter(equipment_user_customer=department)

        # Filter by month/year if specified
        if month and year:
            try:
                month_int = int(month)
                year_int = int(year)
                # Get equipment due in this month
                from calendar import monthrange

                last_day = monthrange(year_int, month_int)[1]
                start_date = datetime(year_int, month_int, 1).date()
                end_date = datetime(year_int, month_int, last_day).date()
                # If requesting PM+CAL, include either PM or CAL due in range
                if pm_cal_flag:
                    from django.db.models import Q

                    qs = qs.filter(
                        Q(equipment_pm_due__gte=start_date, equipment_pm_due__lte=end_date)
                        | Q(equipment_cal_due__gte=start_date, equipment_cal_due__lte=end_date)
                    )
                else:
                    qs = qs.filter(
                        equipment_pm_due__gte=start_date, equipment_pm_due__lte=end_date
                    )
            except (ValueError, TypeError):
                pass

        equipment_list = []

        for eq in qs:
            # Calculate PM and CAL due dates and determine combined status
            pm_due = eq.equipment_pm_due
            cal_due = getattr(eq, 'equipment_cal_due', None)

            # If caller requested pm_cal and neither due exists, skip
            if pm_cal_flag and not pm_due and not cal_due:
                continue

            # Build list of due types present
            due_types = []
            if pm_due:
                due_types.append('PM')
            if cal_due:
                due_types.append('CAL')

            # Choose nearest due date for status/priority calculation
            due_date = None
            if pm_due and cal_due:
                due_date = pm_due if pm_due <= cal_due else cal_due
            else:
                due_date = pm_due or cal_due

            days_until_due = (due_date - today).days if due_date else None

            # Check if already scheduled (future appointments)
            existing_appt = MaintenanceAppointment.objects.filter(
                equipment=eq, scheduled_date__gte=today
            ).first()

            if existing_appt:
                status = "scheduled"
                scheduled_date = existing_appt.scheduled_date.isoformat()
            elif days_until_due is not None and days_until_due < 0:
                status = "overdue"
                scheduled_date = None
            elif days_until_due is not None and days_until_due <= 14:
                status = "due-soon"
                scheduled_date = None
            else:
                status = "upcoming"
                scheduled_date = None

            # Determine priority based on nearest due date
            if days_until_due is None:
                priority = "low"
            elif days_until_due < 0:
                priority = "high"
            elif days_until_due <= 7:
                priority = "high"
            elif days_until_due <= 30:
                priority = "medium"
            else:
                priority = "low"

            # Apply status filter
            if status_filter and status != status_filter:
                continue

            # Calculate frequency display
            pm_fq = getattr(eq, 'equipment_pm_fq', None)
            if isinstance(pm_fq, int) and pm_fq > 0:
                if pm_fq >= 12:
                    years = pm_fq // 12
                    frequency = f"ทุก {years} ปี" if years > 1 else "ทุก 1 ปี"
                else:
                    frequency = f"ทุก {pm_fq} เดือน"
            else:
                frequency = "ไม่ระบุ"

            equipment_list.append(
                {
                    "id": eq.id,
                    "code": eq.equipment_id,
                    "name": eq.equipment_name_TH or eq.equipment_name_EN,
                    "nameEn": eq.equipment_name_EN or "",
                    "department": (
                        eq.equipment_user_customer[:20]
                        if eq.equipment_user_customer
                        else "-"
                    ),  # Shorten
                    "departmentName": eq.equipment_user_customer or "-",
                    "ownerCustomer": eq.equipment_owner_customer or "-",
                    "maintenanceDue": (due_date.isoformat() if due_date else None),
                    "status": status,
                    "priority": priority,
                    "priority_label": {"intime": "ในเวลา", "overtime": "ล่วงเวลา"}.get(
                        priority, priority
                    ),
                    "scheduledDate": scheduled_date,
                    "lastMaintenance": None,  # TODO: Get from WorkOrder history
                    "frequency": frequency,
                    "equipmentType": eq.equipment_type or "-",
                    # Additional equipment details for prefill
                    "brand": getattr(eq, 'equipment_brand', '') or "",
                    "model": getattr(eq, 'equipment_model', '') or "",
                    "serialNumber": getattr(eq, 'equipment_sn', '') or "",
                    "govNumber": getattr(eq, 'equipment_gov', '') or "",
                    # CAL fields and derived due types
                    "calDue": cal_due.isoformat() if cal_due else None,
                    "requires_cal": bool(getattr(eq, 'requires_cal', False)),
                    "dueTypes": due_types,
                    "equipment_pm_due": pm_due.isoformat() if pm_due else None,
                }
            )

        return JsonResponse(
            {
                "success": True,
                "data": equipment_list,
                "count": len(equipment_list),
            }
        )

    except Exception as e:
        logger.error(f"Error in api_equipment_due_list: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


@login_required
@require_http_methods(["POST"])
def api_schedule_equipment(request):
    """API endpoint to schedule maintenance for a single equipment.

    POST data:
    - equipment_id: Equipment ID
    - scheduled_date: Date in YYYY-MM-DD format
    - notes: Optional notes

    Returns success/error message.
    """
    import json
    from datetime import datetime

    from .models import MaintenanceAppointment

    try:
        data = json.loads(request.body)
        equipment_id = data.get("equipment_id")
        scheduled_date = data.get("scheduled_date")
        notes = data.get("notes", "")

        if not equipment_id or not scheduled_date:
            return JsonResponse(
                {"success": False, "error": "กรุณาระบุ equipment_id และ scheduled_date"},
                status=400,
            )

        # Get equipment
        try:
            equipment = Equipment_list.objects.get(id=equipment_id)
        except Equipment_list.DoesNotExist:
            return JsonResponse(
                {"success": False, "error": "ไม่พบเครื่องมือที่ระบุ"}, status=404
            )

        # Parse date
        try:
            sched_date = datetime.strptime(scheduled_date, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse(
                {"success": False, "error": "รูปแบบวันที่ไม่ถูกต้อง (ใช้ YYYY-MM-DD)"},
                status=400,
            )

        # Check if already scheduled (update if exists)
        appt, created = MaintenanceAppointment.objects.get_or_create(
            equipment=equipment,
            scheduled_date__gte=datetime.now().date(),
            defaults={
                "scheduled_date": sched_date,
                "created_by": request.user,
                "notes": notes,
            },
        )

        if not created:
            # Update existing appointment
            appt.scheduled_date = sched_date
            appt.notes = notes
            appt.save()
            message = f"อัพเดทนัดหมายสำหรับ {equipment.equipment_id} เรียบร้อย"
        else:
            message = f"บันทึกนัดหมายสำหรับ {equipment.equipment_id} เรียบร้อย"

        return JsonResponse(
            {
                "success": True,
                "message": message,
                "appointment_id": appt.id,
                "scheduled_date": appt.scheduled_date.isoformat(),
            }
        )

    except json.JSONDecodeError:
        return JsonResponse(
            {"success": False, "error": "ข้อมูล JSON ไม่ถูกต้อง"}, status=400
        )
    except Exception as e:
        logger.error(f"Error in api_schedule_equipment: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


@login_required
@require_http_methods(["POST"])
def api_bulk_schedule_equipment(request):
    """API endpoint to schedule maintenance for multiple equipment at once.

    POST data:
    - equipment_ids: Array of equipment IDs
    - scheduled_date: Date in YYYY-MM-DD format
    - notes: Optional notes

    Returns success count and any errors.
    """
    import json
    from datetime import datetime

    from .models import MaintenanceAppointment

    try:
        data = json.loads(request.body)
        equipment_ids = data.get("equipment_ids", [])
        scheduled_date = data.get("scheduled_date")
        notes = data.get("notes", "")

        if not equipment_ids or not scheduled_date:
            return JsonResponse(
                {"success": False, "error": "กรุณาระบุ equipment_ids และ scheduled_date"},
                status=400,
            )

        # Parse date
        try:
            sched_date = datetime.strptime(scheduled_date, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse(
                {"success": False, "error": "รูปแบบวันที่ไม่ถูกต้อง (ใช้ YYYY-MM-DD)"},
                status=400,
            )

        success_count = 0
        errors = []

        for eq_id in equipment_ids:
            try:
                equipment = Equipment_list.objects.get(id=eq_id)

                # Create or update appointment
                appt, created = MaintenanceAppointment.objects.get_or_create(
                    equipment=equipment,
                    scheduled_date__gte=datetime.now().date(),
                    defaults={
                        "scheduled_date": sched_date,
                        "created_by": request.user,
                        "notes": notes,
                    },
                )

                if not created:
                    appt.scheduled_date = sched_date
                    appt.notes = notes
                    appt.save()

                success_count += 1

            except Equipment_list.DoesNotExist:
                errors.append(f"ไม่พบเครื่องมือ ID: {eq_id}")
            except Exception as e:
                errors.append(f"เกิดข้อผิดพลาดกับ ID {eq_id}: {str(e)}")

        return JsonResponse(
            {
                "success": True,
                "message": f"บันทึกนัดหมายสำเร็จ {success_count} รายการ",
                "success_count": success_count,
                "errors": errors,
            }
        )

    except json.JSONDecodeError:
        return JsonResponse(
            {"success": False, "error": "ข้อมูล JSON ไม่ถูกต้อง"}, status=400
        )
    except Exception as e:
        logger.error(f"Error in api_bulk_schedule_equipment: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


@login_required
@require_http_methods(["POST", "DELETE"])
def api_auto_assign_appointment(request):
    """Auto-assign a technician and create a MaintenanceAppointment for the given
    equipment + date using the same capacity/availability logic as capacity_summary.

    POST body (JSON or form):
      - equipment_id:              Equipment code / equipment_id field
      - date:                      ISO date string  YYYY-MM-DD
      - equipment_type:            (optional) override equipment type for skill matching
      - previous_appointment_id:   (optional) existing appointment to replace / delete
    """
    import json as _json
    from datetime import date as _date
    from django.db import transaction

    # ── DELETE: cancel/delete a pending appointment ──────────────────────────
    if request.method == "DELETE":
        try:
            body = _json.loads(request.body)
        except Exception:
            body = {}
        appt_id = body.get("appointment_id")
        if appt_id:
            from .models import MaintenanceAppointment
            MaintenanceAppointment.objects.filter(
                id=int(appt_id), notes="auto-assign:pending_sr"
            ).delete()
        return JsonResponse({"success": True})

    # ── parse POST body ──────────────────────────────────────────────────────
    try:
        body = _json.loads(request.body)
    except Exception:
        body = request.POST

    equipment_id = (body.get("equipment_id") or "").strip()
    date_str = (body.get("date") or "").strip()
    equipment_type = (body.get("equipment_type") or "").strip()
    prev_appt_id = body.get("previous_appointment_id")

    if not equipment_id or not date_str:
        return JsonResponse({"success": False, "error": "กรุณาระบุรหัสเครื่องมือและวันที่"})

    # ── validate date ────────────────────────────────────────────────────────
    try:
        appt_date = _date.fromisoformat(date_str)
    except (ValueError, TypeError):
        return JsonResponse({"success": False, "error": "รูปแบบวันที่ไม่ถูกต้อง"})

    from .models import (
        Equipment_list,
        MaintenanceAppointment,
        MaintenanceCapacity,
        Technician,
        TechnicianAvailability,
    )

    # ── resolve equipment ────────────────────────────────────────────────────
    eq = (
        Equipment_list.objects.filter(equipment_id=equipment_id).first()
        or Equipment_list.objects.filter(equipment_code=equipment_id).first()
    )
    if not eq:
        return JsonResponse({"success": False, "error": f"ไม่พบเครื่องมือ '{equipment_id}'"})

    eq_type = equipment_type or eq.equipment_type or ""

    # ── cancel / delete previous pending appointment ─────────────────────────
    if prev_appt_id:
        try:
            MaintenanceAppointment.objects.filter(
                id=int(prev_appt_id), notes="auto-assign:pending_sr"
            ).delete()
        except Exception:
            pass

    # ════════════════════════════════════════════════════════════════════════
    #  Capacity-summary logic (mirrors capacity_summary view)
    # ════════════════════════════════════════════════════════════════════════

    # Working statuses — same set as capacity_summary
    WORKING_STATUSES = {"working", "work_regular", "overtime", "work_shift", "work_ot"}

    # Build skill → tech mapping from skills_m (ManyToMany) — same as capacity_summary
    all_techs = list(Technician.objects.filter(active=True).prefetch_related("skills_m"))
    skill_to_techs: dict = {}
    for tech in all_techs:
        for skill in tech.skills_m.all():
            skill_to_techs.setdefault(skill.label, []).append(tech)

    # Find matching skill labels.
    # Try multiple equipment-derived candidates (equipment_type, equipment_code,
    # equipment_name_EN, equipment_name_TH) so that broad category values
    # (e.g. 'Diagnostic') don't block matching when tech skills use specific
    # labels like 'NIBP Monitor'. Use substring match both ways (label in cand
    # or cand in label) — mirrors capacity_summary but with additional fallbacks.
    candidates = []
    if eq_type:
        candidates.append(eq_type)
    # include code and names as additional matching candidates
    try:
        if getattr(eq, "equipment_code", None):
            candidates.append(eq.equipment_code)
    except Exception:
        pass
    try:
        if getattr(eq, "equipment_name_EN", None):
            candidates.append(eq.equipment_name_EN)
    except Exception:
        pass
    try:
        if getattr(eq, "equipment_name_TH", None):
            candidates.append(eq.equipment_name_TH)
    except Exception:
        pass

    eq_lowers = [c.lower() for c in candidates if c]

    matching_labels = []
    if eq_lowers:
        for lbl in skill_to_techs:
            lbl_lower = lbl.lower()
            for cand in eq_lowers:
                if lbl_lower in cand or cand in lbl_lower:
                    matching_labels.append(lbl)
                    break

    # If we still have no matches but an explicit equipment_type was provided,
    # return a helpful error so callers/users can fix data or add skills.
    if not matching_labels and eq_type:
        return JsonResponse(
            {
                "success": False,
                "error": (
                    f"ไม่พบช่างที่มีทักษะ '{eq_type}' ในระบบ"
                    " — กรุณาเพิ่มทักษะให้ช่างหรือติดต่อเจ้าหน้าที่"
                ),
            }
        )

    # Collect skilled techs (deduplicated)
    if matching_labels:
        skilled_ids = {t.id for lbl in matching_labels for t in skill_to_techs[lbl]}
        skilled_techs = [t for t in all_techs if t.id in skilled_ids]
    else:
        # No eq_type provided — consider all active techs
        skilled_techs = all_techs

    if not skilled_techs:
        return JsonResponse({"success": False, "error": "ไม่พบช่างในระบบ"})

    # ── availability check (capacity_summary date_has_records logic) ──────────
    avail_records = {
        av["technician_id"]: av["status"]
        for av in TechnicianAvailability.objects.filter(date=appt_date).values(
            "technician_id", "status"
        )
    }
    # If ANY availability records exist for this date → unknown tech = off.
    # If NO records → planning mode → assume everyone is working.
    date_has_records = bool(avail_records)

    available_techs = []
    for tech in skilled_techs:
        if date_has_records:
            status = avail_records.get(tech.id, "off")
        else:
            status = avail_records.get(tech.id, "working")
        if status in WORKING_STATUSES:
            available_techs.append(tech)

    if not available_techs:
        return JsonResponse(
            {"success": False, "error": "ไม่พบช่างที่ว่างในวันที่เลือก (ตรวจสอบจากตาราง capacity/summary)"}
        )

    # ── effective capacity check (mirrors capacity_summary effective formula) ──
    # avail_cap = sum of per_day_capacity of available techs
    avail_cap = sum(t.per_day_capacity for t in available_techs)

    # Configured capacity override (min of configured vs avail_cap, for each matching label)
    effective_cap = avail_cap
    for lbl in matching_labels:
        cfg = MaintenanceCapacity.objects.filter(
            date=appt_date, equipment_type=lbl
        ).values_list("capacity", flat=True).first()
        if cfg is not None:
            effective_cap = min(effective_cap, cfg)

    # ── Count booked using SAME logic as capacity_summary ────────────────────
    # capacity_summary counts ALL appointments whose equipment type/name
    # matches the skill label via substring match — NOT filtered by assigned_technician.
    # api_available_dates_for_equipment uses equipment__equipment_type exact match.
    # We use exact match here (same as api_available_dates_for_equipment) so
    # remaining counts shown in the date picker == what we enforce here.
    total_booked = MaintenanceAppointment.objects.filter(
        scheduled_date=appt_date,
        equipment__equipment_type=eq_type,   # exact match — same source as date picker API
    ).count()

    if total_booked >= effective_cap:
        return JsonResponse(
            {"success": False, "error": "ความจุของวันนี้เต็มแล้ว กรุณาเลือกวันอื่น (ตรวจสอบที่ capacity/summary)"}
        )

    # ── load-balance: pick tech with least current load who still has capacity ──
    # Per-tech load: count only appointments WITH assigned_technician (for fair distribution)
    per_tech_load: dict = {}
    for tid in MaintenanceAppointment.objects.filter(
        assigned_technician__in=available_techs, scheduled_date=appt_date
    ).values_list("assigned_technician_id", flat=True):
        per_tech_load[tid] = per_tech_load.get(tid, 0) + 1

    best_tech = None
    best_load = None
    for tech in available_techs:
        used = per_tech_load.get(tech.id, 0)
        if used < tech.per_day_capacity:
            if best_load is None or used < best_load:
                best_tech = tech
                best_load = used

    if best_tech is None:
        return JsonResponse(
            {"success": False, "error": "ช่างทุกคนรับงานเต็มแล้วในวันที่เลือก กรุณาเลือกวันอื่น"}
        )

    # ── create appointment ───────────────────────────────────────────────────
    with transaction.atomic():
        appt = MaintenanceAppointment.objects.create(
            equipment=eq,
            scheduled_date=appt_date,
            assigned_technician=best_tech,
            created_by=request.user,
            notes="auto-assign:pending_sr",
        )

    # Remaining = effective_cap − (total_booked + 1 just created)
    remaining = max(0, effective_cap - total_booked - 1)

    return JsonResponse(
        {
            "success": True,
            "appointment_id": appt.id,
            "technician_name": best_tech.name,
            "technician_id": best_tech.id,
            "remaining": remaining,
            "scheduled_date": str(appt_date),
            # diagnostics (helpful for debugging)
            "effective_capacity": effective_cap,
            "booked_before": total_booked,
        }
    )


@login_required
@require_http_methods(["POST"])
def api_create_service_request(request):
    """API endpoint to create service request for equipment maintenance.

    POST data:
    - equipment_id: Equipment ID
    - requested_date: Date in YYYY-MM-DD format (when equipment should be sent for maintenance)
    - notes: Optional notes

    Returns success/error message with service request ID.
    """
    import json
    from datetime import datetime

    from .models import MaintenanceAppointment, ServiceRequest

    try:
        data = json.loads(request.body)
        equipment_id = data.get("equipment_id")
        requested_date = data.get("requested_date")
        notes = data.get("notes", "")

        if not equipment_id or not requested_date:
            return JsonResponse(
                {"success": False, "error": "กรุณาระบุ equipment_id และ requested_date"},
                status=400,
            )

        # Get equipment
        try:
            equipment = Equipment_list.objects.get(id=equipment_id)
        except Equipment_list.DoesNotExist:
            return JsonResponse(
                {"success": False, "error": "ไม่พบเครื่องมือที่ระบุ"}, status=404
            )

        # Parse date
        try:
            req_date = datetime.strptime(requested_date, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse(
                {"success": False, "error": "รูปแบบวันที่ไม่ถูกต้อง (ใช้ YYYY-MM-DD)"},
                status=400,
            )

        # Create service request
        service_request = ServiceRequest.objects.create(
            title=f"บำรุงรักษาเครื่อง {equipment.equipment_id} - {equipment.equipment_name}",
            description=f"คำร้องขอบำรุงรักษาตามกำหนด PM\nกำหนดบำรุงรักษา: {equipment.equipment_pm_due}\nวันที่ต้องการนำเครื่องเข้าบำรุงรักษา: {req_date}",
            equipment=equipment,
            requested_by=request.user,
            status="new",
            notes=notes,
        )

        # Also create maintenance appointment for tracking
        MaintenanceAppointment.objects.get_or_create(
            equipment=equipment,
            scheduled_date__gte=datetime.now().date(),
            defaults={
                "scheduled_date": req_date,
                "created_by": request.user,
                "notes": f"Service Request #{service_request.id}",
            },
        )

        message = f"สร้างคำร้องขอบำรุงรักษา SR#{service_request.id} สำหรับ {equipment.equipment_id} เรียบร้อย"

        return JsonResponse(
            {
                "success": True,
                "message": message,
                "service_request_id": service_request.id,
                "requested_date": req_date.isoformat(),
            }
        )

    except json.JSONDecodeError:
        return JsonResponse(
            {"success": False, "error": "ข้อมูล JSON ไม่ถูกต้อง"}, status=400
        )
    except Exception as e:
        logger.error(f"Error in api_create_service_request: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


@login_required
@require_http_methods(["POST"])
def api_bulk_create_service_request(request):
    """API endpoint to create service requests for multiple equipment at once.

    POST data:
    - equipment_ids: Array of equipment IDs
    - requested_date: Date in YYYY-MM-DD format
    - notes: Optional notes

    Returns success count and any errors.
    """
    import json
    from datetime import datetime

    from .models import MaintenanceAppointment, ServiceRequest

    try:
        data = json.loads(request.body)
        equipment_ids = data.get("equipment_ids", [])
        requested_date = data.get("requested_date")
        notes = data.get("notes", "")

        if not equipment_ids or not requested_date:
            return JsonResponse(
                {"success": False, "error": "กรุณาระบุ equipment_ids และ requested_date"},
                status=400,
            )

        # Parse date
        try:
            req_date = datetime.strptime(requested_date, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse(
                {"success": False, "error": "รูปแบบวันที่ไม่ถูกต้อง (ใช้ YYYY-MM-DD)"},
                status=400,
            )

        success_count = 0
        errors = []

        for eq_id in equipment_ids:
            try:
                equipment = Equipment_list.objects.get(id=eq_id)

                # Create service request
                service_request = ServiceRequest.objects.create(
                    title=f"บำรุงรักษาเครื่อง {equipment.equipment_id} - {equipment.equipment_name}",
                    description=f"คำร้องขอบำรุงรักษาตามกำหนด PM\nกำหนดบำรุงรักษา: {equipment.equipment_pm_due}\nวันที่ต้องการนำเครื่องเข้าบำรุงรักษา: {req_date}",
                    equipment=equipment,
                    requested_by=request.user,
                    status="new",
                    notes=notes,
                )

                # Also create maintenance appointment
                MaintenanceAppointment.objects.get_or_create(
                    equipment=equipment,
                    scheduled_date__gte=datetime.now().date(),
                    defaults={
                        "scheduled_date": req_date,
                        "created_by": request.user,
                        "notes": f"Service Request #{service_request.id}",
                    },
                )

                success_count += 1

            except Equipment_list.DoesNotExist:
                errors.append(f"ไม่พบเครื่องมือ ID: {eq_id}")
            except Exception as e:
                errors.append(f"เกิดข้อผิดพลาดกับ ID {eq_id}: {str(e)}")

        return JsonResponse(
            {
                "success": True,
                "message": f"สร้างคำร้องขอบำรุงรักษาสำเร็จ {success_count} รายการ",
                "success_count": success_count,
                "errors": errors,
            }
        )

    except json.JSONDecodeError:
        return JsonResponse(
            {"success": False, "error": "ข้อมูล JSON ไม่ถูกต้อง"}, status=400
        )
    except Exception as e:
        logger.error(f"Error in api_bulk_create_service_request: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


@login_required
@require_http_methods(["POST"])
def api_service_requests_by_equipment(request):
    """Return current user's non-closed ServiceRequests grouped by equipment_id.

    POST JSON: { "equipment_ids": [1,2,3] }

    Response: { success: true, results: { "1": [ {id, status, requested_at}, ... ], ... } }
    """
    import json
    try:
        data = json.loads(request.body)
        equipment_ids = data.get('equipment_ids', []) if isinstance(data, dict) else []
        if not equipment_ids:
            return JsonResponse({"success": True, "results": {}})

        from .models import ServiceRequest

        qs = (
            ServiceRequest.objects
            .filter(requested_by=request.user, equipment_id__in=equipment_ids)
            .exclude(status='closed')
            .order_by('-requested_at')
        )

        results = {}
        for sr in qs:
            key = str(sr.equipment_id)
            results.setdefault(key, []).append({
                'id': sr.id,
                'status': sr.status,
                'requested_at': sr.requested_at.isoformat() if getattr(sr, 'requested_at', None) else None
            })

        return JsonResponse({"success": True, "results": results})
    except json.JSONDecodeError:
        return JsonResponse({"success": False, "error": "Invalid JSON"}, status=400)
    except Exception as e:
        logger.error(f"Error in api_service_requests_by_equipment: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e)}, status=500)


# Master Category management (in-app UI for staff) ----------------------------
@staff_member_required
def category_list(request):
    """List all MasterCategory objects with links to edit."""
    from .models import MasterCategory

    categories = MasterCategory.objects.all().order_by("label")
    return render(request, "master/category_list.html", {"categories": categories})


@staff_member_required
@require_http_methods(["GET", "POST"])
def category_create(request):
    """Create a new MasterCategory with inline field definitions (formset)."""
    from django import forms
    from django.forms import inlineformset_factory

    from .models import MasterCategory, MasterFieldDefinition

    class CategoryForm(forms.ModelForm):
        class Meta:
            model = MasterCategory
            fields = ("key", "label", "description", "display_template")
            widgets = {
                "key": forms.TextInput(attrs={"class": "form-control"}),
                "label": forms.TextInput(attrs={"class": "form-control"}),
                "description": forms.Textarea(
                    attrs={"class": "form-control", "rows": 3}
                ),
                "display_template": forms.Textarea(
                    attrs={"class": "form-control", "rows": 3}
                ),
            }

    FieldDefFormSet = inlineformset_factory(
        MasterCategory,
        MasterFieldDefinition,
        fields=(
            "name",
            "label",
            "field_type",
            "required",
            "choices",
            "default_value",
            "order",
            "help_text",
            "visible_in_preview",
        ),
        extra=3,
        can_delete=True,
        widgets={
            "name": forms.TextInput(
                attrs={
                    "class": "form-control form-control-sm",
                    "placeholder": "machine_key",
                }
            ),
            "label": forms.TextInput(
                attrs={
                    "class": "form-control form-control-sm",
                    "placeholder": "แสดงใน UI",
                }
            ),
            "field_type": forms.Select(attrs={"class": "form-select form-select-sm"}),
            "required": forms.CheckboxInput(attrs={"class": "form-check-input"}),
            "choices": forms.Textarea(
                attrs={
                    "class": "form-control form-control-sm",
                    "rows": 2,
                    "placeholder": "JSON array",
                }
            ),
            "default_value": forms.TextInput(
                attrs={"class": "form-control form-control-sm"}
            ),
            "order": forms.NumberInput(
                attrs={"class": "form-control form-control-sm", "value": 0}
            ),
            "help_text": forms.TextInput(
                attrs={"class": "form-control form-control-sm"}
            ),
            "visible_in_preview": forms.CheckboxInput(
                attrs={"class": "form-check-input"}
            ),
        },
    )

    if request.method == "POST":
        form = CategoryForm(request.POST)
        formset = FieldDefFormSet(
            request.POST, instance=form.instance if form.is_valid() else None
        )
        if form.is_valid() and formset.is_valid():
            cat = form.save()
            formset.instance = cat
            formset.save()
            messages.success(request, f"สร้างหมวด {cat.label} เรียบร้อยแล้ว")
            return redirect("category_list")
    else:
        form = CategoryForm()
        formset = FieldDefFormSet()

    # Provide master categories and items (for the template's "load choices" UI).
    try:
        from .models import MasterCategory, MasterItem

        # Build category list from existing MasterItem rows (only categories that have items).
        # Use MasterCategory.label when available, otherwise fall back to the raw category key.
        cat_keys = list(
            MasterItem.objects.filter(active=True)
            .order_by("category")
            .values_list("category", flat=True)
            .distinct()
        )
        master_categories = []
        for ck in cat_keys:
            try:
                mc = MasterCategory.objects.filter(key=ck).first()
                label = mc.label if mc else ck
            except Exception:
                label = ck
            master_categories.append({"key": ck, "label": label})

        qs = MasterItem.objects.filter(active=True).order_by("category", "order", "label")
        master_items_by_category = {}
        for m in qs:
            cat = m.category or ""
            master_items_by_category.setdefault(cat, [])
            master_items_by_category[cat].append(
                {
                    "value": (m.code or str(m.id)),
                    "code": (m.code or ""),
                    "key": m.id,
                    "label": (m.label or ""),
                    "name": (m.label or ""),
                }
            )
    except Exception:
        master_categories = []
        master_items_by_category = {}

    return render(
        request,
        "master/category_form.html",
        {
            "form": form,
            "formset": formset,
            "create": True,
            "master_categories": master_categories,
            "master_items_by_category": master_items_by_category,
        },
    )


@staff_member_required
@require_http_methods(["GET", "POST"])
def category_edit(request, cat_id):
    """Edit an existing MasterCategory and its field definitions (inline formset)."""
    from django import forms
    from django.forms import inlineformset_factory

    from .models import MasterCategory, MasterFieldDefinition

    cat = get_object_or_404(MasterCategory, pk=cat_id)

    class CategoryForm(forms.ModelForm):
        class Meta:
            model = MasterCategory
            fields = ("key", "label", "description", "display_template")
            widgets = {
                "key": forms.TextInput(attrs={"class": "form-control"}),
                "label": forms.TextInput(attrs={"class": "form-control"}),
                "description": forms.Textarea(
                    attrs={"class": "form-control", "rows": 3}
                ),
                "display_template": forms.Textarea(
                    attrs={"class": "form-control", "rows": 3}
                ),
            }

    FieldDefFormSet = inlineformset_factory(
        MasterCategory,
        MasterFieldDefinition,
        fields=(
            "name",
            "label",
            "field_type",
            "required",
            "choices",
            "default_value",
            "order",
            "help_text",
            "visible_in_preview",
        ),
        extra=1,
        can_delete=True,
        widgets={
            "name": forms.TextInput(
                attrs={
                    "class": "form-control form-control-sm",
                    "placeholder": "machine_key",
                }
            ),
            "label": forms.TextInput(
                attrs={
                    "class": "form-control form-control-sm",
                    "placeholder": "แสดงใน UI",
                }
            ),
            "field_type": forms.Select(attrs={"class": "form-select form-select-sm"}),
            "required": forms.CheckboxInput(attrs={"class": "form-check-input"}),
            "choices": forms.Textarea(
                attrs={
                    "class": "form-control form-control-sm",
                    "rows": 2,
                    "placeholder": "JSON array",
                }
            ),
            "default_value": forms.TextInput(
                attrs={"class": "form-control form-control-sm"}
            ),
            "order": forms.NumberInput(attrs={"class": "form-control form-control-sm"}),
            "help_text": forms.TextInput(
                attrs={"class": "form-control form-control-sm"}
            ),
            "visible_in_preview": forms.CheckboxInput(
                attrs={"class": "form-check-input"}
            ),
        },
    )

    if request.method == "POST":
        form = CategoryForm(request.POST, instance=cat)
        formset = FieldDefFormSet(request.POST, instance=cat)
        if form.is_valid() and formset.is_valid():
            form.save()
            formset.save()
            messages.success(request, f"แก้ไขหมวด {cat.label} เรียบร้อยแล้ว")
            return redirect("category_list")
    else:
        form = CategoryForm(instance=cat)
        formset = FieldDefFormSet(instance=cat)

    # Provide master categories and items so the template can populate the choices picker
    try:
        from .models import MasterCategory, MasterItem

        # Build category list from existing MasterItem rows (only categories that have items).
        # Use MasterCategory.label when available, otherwise fall back to the raw category key.
        cat_keys = list(
            MasterItem.objects.filter(active=True)
            .order_by("category")
            .values_list("category", flat=True)
            .distinct()
        )
        master_categories = []
        for ck in cat_keys:
            try:
                mc = MasterCategory.objects.filter(key=ck).first()
                label = mc.label if mc else ck
            except Exception:
                label = ck
            master_categories.append({"key": ck, "label": label})

        qs = MasterItem.objects.filter(active=True).order_by("category", "order", "label")
        master_items_by_category = {}
        for m in qs:
            catk = m.category or ""
            master_items_by_category.setdefault(catk, [])
            master_items_by_category[catk].append(
                {
                    "value": (m.code or str(m.id)),
                    "code": (m.code or ""),
                    "key": m.id,
                    "label": (m.label or ""),
                    "name": (m.label or ""),
                }
            )
    except Exception:
        master_categories = []
        master_items_by_category = {}

    return render(
        request,
        "master/category_form.html",
        {
            "form": form,
            "formset": formset,
            "create": False,
            "category": cat,
            "master_categories": master_categories,
            "master_items_by_category": master_items_by_category,
        },
    )


@staff_member_required
@require_http_methods(["POST"])
def category_delete(request, cat_id):
    """Delete a MasterCategory (and cascade delete its field definitions)."""
    from .models import MasterCategory

    cat = get_object_or_404(MasterCategory, pk=cat_id)
    label = cat.label
    cat.delete()
    messages.success(request, f"ลบหมวด {label} เรียบร้อยแล้ว")
    return redirect("category_list")


# ══════════════════════════════════════════════════════════════════════════════
# ทะเบียนอะไหล่คงคลัง (Spare Parts Inventory)
# ══════════════════════════════════════════════════════════════════════════════
from .models import SparePart


@login_required
def spare_part_list(request):
    """List all spare parts with search, filter, and low-stock alerts."""
    from django.db.models import Q, F

    q = request.GET.get("q", "").strip()
    status_filter = request.GET.get("status", "")  # all / low / active / inactive

    qs = SparePart.objects.select_related("category", "created_by").all()

    if q:
        qs = qs.filter(
            Q(code__icontains=q)
            | Q(name__icontains=q)
            | Q(supplier__icontains=q)
            | Q(location__icontains=q)
        )

    if status_filter == "low":
        qs = qs.filter(quantity__lte=F("min_quantity"))
    elif status_filter == "active":
        qs = qs.filter(active=True)
    elif status_filter == "inactive":
        qs = qs.filter(active=False)

    total_items = qs.count()
    low_stock_count = SparePart.objects.filter(
        quantity__lte=F("min_quantity"), active=True
    ).count()
    total_value = sum(
        (p.quantity * p.unit_cost) for p in qs if p.unit_cost
    )
    # prepare master lists for selects (manufacturers -> models)
    from .models import MasterItem

    def _ml(cat):
        qs_m = list(
            MasterItem.objects.filter(category=cat, active=True).order_by("order", "label")
        )
        if not qs_m:
            return []
        label_map = {m.id: m.label for m in qs_m}
        parent_map = {m.id: (m.parent_id if m.parent_id else None) for m in qs_m}

        def build_ancestors(mid):
            chain = []
            seen = set()
            cur = parent_map.get(mid)
            while cur and cur not in seen:
                seen.add(cur)
                lbl = label_map.get(cur)
                if lbl:
                    chain.append(lbl)
                cur = parent_map.get(cur)
            chain.reverse()
            return " > ".join(chain)

        out = []
        for m in qs_m:
            out.append({"id": m.id, "label": m.label, "ancestors": build_ancestors(m.id)})
        return out

    # Categories for the spare part category dropdown
    categories = list(
        MasterItem.objects.filter(category="spare_part_categories", active=True)
        .order_by("order", "label")
    )

    return render(
        request,
        "spare_parts/spare_part_list.html",
        {
            "parts": qs,
            "q": q,
            "status_filter": status_filter,
            "total_items": total_items,
            "low_stock_count": low_stock_count,
            "total_value": total_value,
            "master_manufacturers": _ml("manufacturers"),
            "master_equipments": _ml("equipments"),
            "categories": categories,
            "form_values": {},
            "errors": {},
        },
    )


@login_required
def spare_part_create(request):
    """Create a new spare part (AJAX modal)."""
    if request.method == "POST":
        code = request.POST.get("code", "").strip()
        name = request.POST.get("name", "").strip()

        if not code or not name:
            return JsonResponse({"ok": False, "error": "กรุณากรอกรหัสและชื่ออะไหล่"})

        if SparePart.objects.filter(code=code).exists():
            return JsonResponse({"ok": False, "error": f"รหัส {code} มีในระบบแล้ว"})

        part = SparePart(
            code=code,
            name=name,
            description=request.POST.get("description", ""),
            unit=request.POST.get("unit", "ชิ้น"),
            quantity=int(request.POST.get("quantity", 0) or 0),
            min_quantity=int(request.POST.get("min_quantity", 0) or 0),
            location=request.POST.get("location", ""),
            supplier=request.POST.get("supplier", ""),
            notes=request.POST.get("notes", ""),
            active=True,
            created_by=request.user,
            updated_by=request.user,
        )
        cost = request.POST.get("unit_cost", "")
        if cost:
            from decimal import Decimal, InvalidOperation
            try:
                part.unit_cost = Decimal(cost)
            except (InvalidOperation, ValueError):
                pass
        cat_id = request.POST.get("category", "")
        if cat_id:
            part.category_id = int(cat_id)

        # equipment_meta from multi-select dropdowns
        import json as _json
        part.equipment_meta = {
            "equipment_names": _json.loads(request.POST.get("equipment_names", "[]")),
            "equipment_brands": _json.loads(request.POST.get("equipment_brands", "[]")),
            "equipment_models": _json.loads(request.POST.get("equipment_models", "[]")),
        }

        part.save()

        # M2M compatible equipment
        eq_ids = request.POST.getlist("compatible_equipment")
        if eq_ids:
            part.compatible_equipment.set(eq_ids)

        return JsonResponse({"ok": True, "id": part.id, "code": part.code})

    return JsonResponse({"ok": False, "error": "Method not allowed"}, status=405)


@login_required
def spare_part_update(request, pk):
    """Update an existing spare part (AJAX modal)."""
    part = get_object_or_404(SparePart, pk=pk)

    if request.method == "GET":
        meta = part.equipment_meta or {}
        data = {
            "id": part.id,
            "code": part.code,
            "name": part.name,
            "description": part.description,
            "category": part.category_id,
            "unit": part.unit,
            "quantity": part.quantity,
            "min_quantity": part.min_quantity,
            "unit_cost": str(part.unit_cost) if part.unit_cost else "",
            "location": part.location,
            "supplier": part.supplier,
            "notes": part.notes,
            "active": part.active,
            "compatible_equipment": list(
                part.compatible_equipment.values_list("id", flat=True)
            ),
            "equipment_names": meta.get("equipment_names", []),
            "equipment_brands": meta.get("equipment_brands", []),
            "equipment_models": meta.get("equipment_models", []),
        }
        return JsonResponse({"ok": True, "data": data})

    if request.method == "POST":
        part.code = request.POST.get("code", part.code).strip()
        part.name = request.POST.get("name", part.name).strip()

        if not part.code or not part.name:
            return JsonResponse({"ok": False, "error": "กรุณากรอกรหัสและชื่ออะไหล่"})

        # Check unique code (exclude self)
        if SparePart.objects.filter(code=part.code).exclude(pk=pk).exists():
            return JsonResponse({"ok": False, "error": f"รหัส {part.code} มีในระบบแล้ว"})

        part.description = request.POST.get("description", "")
        part.unit = request.POST.get("unit", "ชิ้น")
        part.quantity = int(request.POST.get("quantity", 0) or 0)
        part.min_quantity = int(request.POST.get("min_quantity", 0) or 0)
        part.location = request.POST.get("location", "")
        part.supplier = request.POST.get("supplier", "")
        part.notes = request.POST.get("notes", "")
        part.active = request.POST.get("active") in ("true", "True", "1", "on")
        part.updated_by = request.user

        cost = request.POST.get("unit_cost", "")
        if cost:
            from decimal import Decimal, InvalidOperation
            try:
                part.unit_cost = Decimal(cost)
            except (InvalidOperation, ValueError):
                pass
        else:
            part.unit_cost = None

        cat_id = request.POST.get("category", "")
        part.category_id = int(cat_id) if cat_id else None

        # equipment_meta from multi-select dropdowns
        import json as _json
        part.equipment_meta = {
            "equipment_names": _json.loads(request.POST.get("equipment_names", "[]")),
            "equipment_brands": _json.loads(request.POST.get("equipment_brands", "[]")),
            "equipment_models": _json.loads(request.POST.get("equipment_models", "[]")),
        }

        part.save()

        eq_ids = request.POST.getlist("compatible_equipment")
        part.compatible_equipment.set(eq_ids)

        return JsonResponse({"ok": True})

    return JsonResponse({"ok": False, "error": "Method not allowed"}, status=405)


@login_required
@require_http_methods(["POST"])
def spare_part_delete(request, pk):
    """Delete a spare part."""
    part = get_object_or_404(SparePart, pk=pk)
    part.delete()
    return JsonResponse({"ok": True})


@login_required
def spare_part_adjust_stock(request, pk):
    """Adjust stock quantity (เบิก/รับเข้า)."""
    part = get_object_or_404(SparePart, pk=pk)

    if request.method == "POST":
        action = request.POST.get("action")  # "add" or "withdraw"
        try:
            amount = int(request.POST.get("amount", 0))
        except (ValueError, TypeError):
            return JsonResponse({"ok": False, "error": "จำนวนไม่ถูกต้อง"})

        if amount <= 0:
            return JsonResponse({"ok": False, "error": "จำนวนต้องมากกว่า 0"})

        if action == "withdraw":
            if amount > part.quantity:
                return JsonResponse({"ok": False, "error": "จำนวนคงเหลือไม่เพียงพอ"})
            part.quantity -= amount
        elif action == "add":
            part.quantity += amount
        else:
            return JsonResponse({"ok": False, "error": "action ไม่ถูกต้อง"})

        part.updated_by = request.user
        part.save(update_fields=["quantity", "updated_by", "updated_at"])
        return JsonResponse({"ok": True, "quantity": part.quantity})

    return JsonResponse({"ok": False, "error": "Method not allowed"}, status=405)


@login_required
def api_spare_parts_search(request):
    """API: ค้นหาอะไหล่สำหรับเลือกเบิกในใบงาน"""
    from django.db.models import Q

    q = request.GET.get("q", "").strip()
    equipment_id = request.GET.get("equipment_id", "").strip()

    qs = SparePart.objects.filter(active=True)

    if q:
        qs = qs.filter(Q(code__icontains=q) | Q(name__icontains=q))

    # Prioritise parts compatible with the WO equipment
    if equipment_id:
        try:
            qs_compat = qs.filter(compatible_equipment__id=int(equipment_id))
            qs_other = qs.exclude(compatible_equipment__id=int(equipment_id))
            results = list(qs_compat[:30]) + list(qs_other[:20])
        except (ValueError, TypeError):
            results = list(qs[:50])
    else:
        results = list(qs[:50])

    data = [
        {
            "id": sp.id,
            "code": sp.code,
            "name": sp.name,
            "unit": sp.unit,
            "quantity": sp.quantity,
            "unit_cost": str(sp.unit_cost or 0),
            "location": sp.location,
        }
        for sp in results
    ]
    return JsonResponse({"results": data})


@login_required
def api_equipment_by_date(request):
    """API endpoint to get equipment scheduled for maintenance on a specific date.
    
    Query parameters:
    - due_date: Date in YYYY-MM-DD format
    
    Returns JSON with equipment list for that date.
    """
    from datetime import datetime
    
    from django.utils import timezone
    
    try:
        due_date_str = request.GET.get("due_date")
        
        if not due_date_str:
            return JsonResponse({"data": []}, safe=False)
        
        # Parse the date
        try:
            due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse({"data": []}, safe=False)
        
        # Get equipment due on this date
        qs = Equipment_list.objects.filter(
            requires_pm=True,
            equipment_pm_due=due_date
        )
        
        equipment_list = []
        total_hours = 0
        # try to load default durations for task types
        from .models import MaintenanceTaskDuration
        default_durations = {d.task_type: float(d.hours) for d in MaintenanceTaskDuration.objects.all()}

        for eq in qs:
            pm_due = eq.equipment_pm_due
            today = timezone.localtime().date()
            days_until_due = (pm_due - today).days
            
            # Determine status
            if days_until_due < 0:
                status = "overdue"
            elif days_until_due <= 14:
                status = "due-soon"
            else:
                status = "upcoming"
            
            # Calculate frequency display
            pm_fq = eq.equipment_pm_fq
            if pm_fq >= 12:
                frequency = f"ทุก {pm_fq // 12} ปี" if pm_fq // 12 > 1 else "ทุก 1 ปี"
            else:
                frequency = f"ทุก {pm_fq} เดือน"
            
            equipment_list.append({
                "id": eq.id,
                "code": eq.equipment_id,
                "name": eq.equipment_name_TH or eq.equipment_name_EN,
                "departmentName": eq.equipment_user_customer or "-",
                "maintenanceDue": pm_due.isoformat() if pm_due else None,
                "status": status,
                "frequency": frequency,
            })
            # determine estimated hours for this equipment
            # Priority: Equipment_list.estimated_hours -> default by task type
            estimated = None
            if getattr(eq, 'estimated_hours', None) is not None:
                try:
                    estimated = float(eq.estimated_hours)
                except Exception:
                    estimated = None

            # determine task type: prefer calibration if requires_cal True, else maintenance
            task_type = 'calibration' if getattr(eq, 'requires_cal', False) else 'maintenance'
            if estimated is None:
                estimated = default_durations.get(task_type, 1.0)

            # attach estimatedHours to item and accumulate
            equipment_list[-1]['estimatedHours'] = estimated
            total_hours += float(estimated)

        result = {
            'items': equipment_list,
            'totalHours': round(total_hours, 2),
        }

        return JsonResponse(result, safe=False)
    
    except Exception as e:
        logger.error(f"Error in api_equipment_by_date: {str(e)}", exc_info=True)
        return JsonResponse({"data": []}, safe=False)


@login_required
def api_stations_availability(request):
    """API endpoint to get station availability for a specific date.
    
    Query parameters:
    - date: Date in YYYY-MM-DD format
    
    Returns JSON with station availability information.
    """
    from datetime import datetime
    
    from .models import MaintenanceCapacity, MaintenanceAppointment
    
    try:
        date_str = request.GET.get("date")
        
        if not date_str:
            return JsonResponse([], safe=False)
        
        # Parse the date
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse([], safe=False)
        
        # Get all stations (MaintenanceCapacity) for this date
        stations = MaintenanceCapacity.objects.filter(date=target_date)
        
        if not stations.exists():
            # Return empty if no capacity records exist for this date
            return JsonResponse([], safe=False)
        
        stations_list = []
        for station in stations:
            # Count appointments on this date
            appointments_count = MaintenanceAppointment.objects.filter(
                scheduled_date=target_date
            ).count()
            
            capacity = station.capacity or 0
            current_load = appointments_count
            available = current_load < capacity
            
            stations_list.append({
                "id": station.id,
                "name": f"{station.equipment_type} - ความจุ {capacity}",
                "capacity": capacity,
                "currentLoad": current_load,
                "available": available,
                "note": "",
            })
        
        return JsonResponse(stations_list, safe=False)
    
    except Exception as e:
        logger.error(f"Error in api_stations_availability: {str(e)}", exc_info=True)
        return JsonResponse([], safe=False)


@login_required
def api_technicians_availability(request):
    """API endpoint to get technician availability for a specific date.
    
    Query parameters:
    - date: Date in YYYY-MM-DD format
    
    Returns JSON with technician availability information.
    """
    from datetime import datetime
    
    from .models import Technician, TechnicianAvailability, WorkOrder
    
    try:
        date_str = request.GET.get("date")
        
        if not date_str:
            return JsonResponse([], safe=False)
        
        # Parse the date
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse([], safe=False)
        
        # Get all active technicians
        technicians = Technician.objects.filter(active=True)
        
        technicians_list = []
        for tech in technicians:
            # Check availability for this date
            availability = TechnicianAvailability.objects.filter(
                technician=tech,
                date=target_date
            ).first()
            
            is_available = availability.status == 'working' if availability else True
            
            # Count current tasks/work orders on this date
            current_tasks = 0
            if tech.user:
                # Use planned_start which is a DateTime field
                current_tasks = WorkOrder.objects.filter(
                    assigned_to=tech.user,
                    planned_start__date=target_date,
                    status__in=['assigned', 'in_progress']
                ).count()
            
            # Get skills
            skills_list = list(tech.skills_m.values_list('label', flat=True)) if tech.skills_m.exists() else []
            skills_str = ", ".join(skills_list) if skills_list else (tech.skills or "-")
            
            technicians_list.append({
                "id": tech.id,
                "name": tech.name,
                "skills": skills_str,
                "currentTasks": current_tasks,
                "capacity": tech.per_day_capacity,
                "available": is_available and current_tasks < tech.per_day_capacity,
                # availability_status: use the raw TechnicianAvailability.status when present
                # if no availability record exists, return null so frontend can distinguish
                # between explicit 'working' records and implicit/default values
                "availability_status": availability.status if availability else None,
                "note": availability.note if availability else "",
            })
        
        return JsonResponse(technicians_list, safe=False)

    except Exception as e:
        logger.error(f"Error in api_technicians_list: {str(e)}", exc_info=True)
        return JsonResponse([], safe=False)

    except Exception as e:
        logger.error(f"Error in api_technicians_availability: {str(e)}", exc_info=True)
        return JsonResponse([], safe=False)



@login_required
@require_http_methods(["POST"])
def api_update_equipment_estimated_hours(request):
    """API to update Equipment_list.estimated_hours for a given equipment id.

    Expects JSON body: { equipment_id: <id>, estimated_hours: <number> }
    """
    import json
    from .models import Equipment_list
    try:
        data = json.loads(request.body.decode('utf-8'))
        eq_id = data.get('equipment_id')
        hours = data.get('estimated_hours')

        if eq_id is None:
            return JsonResponse({'success': False, 'error': 'equipment_id required'}, status=400)

        eq = Equipment_list.objects.filter(id=eq_id).first()
        if not eq:
            return JsonResponse({'success': False, 'error': 'equipment not found'}, status=404)

        # allow null to clear value
        if hours is None:
            eq.estimated_hours = None
        else:
            try:
                eq.estimated_hours = float(hours)
            except Exception:
                return JsonResponse({'success': False, 'error': 'invalid hours value'}, status=400)

        eq.save()
        return JsonResponse({'success': True, 'message': 'estimated hours updated'})

    except Exception as e:
        logger.error(f"Error in api_update_equipment_estimated_hours: {str(e)}", exc_info=True)
        return JsonResponse({'success': False, 'error': str(e)}, status=500)


@login_required
def api_technicians_list(request):
    """API endpoint to get all active technicians.
    
    Returns JSON with technician information.
    """
    from .models import Technician
    
    try:
        technicians = Technician.objects.filter(active=True)
        
        technicians_list = []
        for tech in technicians:
            # Get skills
            skills_list = list(tech.skills_m.values_list('label', flat=True)) if tech.skills_m.exists() else []
            skills_str = ", ".join(skills_list) if skills_list else (tech.skills or "-")
            
            technicians_list.append({
                "id": tech.id,
                "name": tech.name,
                "skills": skills_str,
                "capacity": tech.per_day_capacity,
                "phone": tech.phone or "",
                "canWorkShifts": tech.can_work_shifts,
            })
        
        return JsonResponse(technicians_list, safe=False)
    except Exception as e:
        logger.error(f"Error in api_technicians_list: {str(e)}", exc_info=True)
        return JsonResponse([], safe=False)


@login_required
def service_request_detail(request, sr_id):
        """Display and update a ServiceRequest (supports POST 'save' action).

        This view intentionally mirrors the template expectations: returns context
        keys `sr`, `attachments`, and `logs`. A POST with `action=save` will
        persist the hidden `title` input and notes to the DB and record a log.
        """
        from .models import ServiceRequest, ServiceRequestAttachment, ServiceRequestLog

        sr = get_object_or_404(ServiceRequest, pk=sr_id)

        if request.method == "POST":
            action = request.POST.get("action")

            # ── create_wo: สร้างใบงานจาก SR โดยตรง ──
            if action == "create_wo" and request.user.has_perm("cmms.add_workorder"):
                from cmms.models import WorkOrder, WorkOrderLog
                from django.contrib.auth import get_user_model
                User = get_user_model()

                # ป้องกันสร้างซ้ำ
                if sr.converted_to_id:
                    messages.info(request, f"คำขอนี้ถูกสร้างใบงานแล้ว (WO#{sr.converted_to_id})")
                    return redirect("workorder_detail", wo_id=sr.converted_to_id)

                priority = request.POST.get("wo_priority", "intime")
                if priority not in ("intime", "overtime"):
                    priority = "intime"

                wo = WorkOrder.objects.create(
                    title=sr.title,
                    description=sr.description,
                    equipment=sr.equipment,
                    priority=priority,
                    reported_by=request.user,
                    accepted_by=request.user,
                    accepted_at=timezone.now(),
                )

                # มอบหมายช่าง (ถ้าเลือก)
                assigned_to_id = request.POST.get("wo_assigned_to")
                if assigned_to_id:
                    try:
                        wo.assigned_to = User.objects.get(id=int(assigned_to_id))
                        wo.status = "assigned"
                        wo.save()
                    except Exception:
                        pass

                # อัปเดต SR
                sr.converted_to = wo
                sr.status = "converted"
                sr.save()

                ServiceRequestLog.objects.create(
                    request=sr,
                    action="รับคำขอบริการ",
                    actor=request.user,
                    note=f"สร้างใบงาน WO#{wo.id}",
                )
                WorkOrderLog.objects.create(
                    workorder=wo,
                    action="created_from_request",
                    actor=request.user,
                    note=f"From SR#{sr.id}",
                )

                messages.success(request, f"สร้างใบงาน WO#{wo.id} จากคำขอ SR#{sr.id} สำเร็จ")
                return redirect("workorder_detail", wo_id=wo.id)

            # Allow officers to save back edited title/notes
            if action == "save" and request.user.has_perm("cmms.change_servicerequest"):
                new_title = (request.POST.get("title") or "").strip()
                new_notes = (request.POST.get("notes") or "").strip()

                # capture originals for logging
                old_title = sr.title or ""
                old_notes = sr.notes or ""

                changed = False
                if new_title != old_title:
                    sr.title = new_title
                    changed = True
                if new_notes != old_notes:
                    sr.notes = new_notes
                    changed = True

                if changed:
                    sr.save()
                    # build a descriptive note showing before -> after for changed fields
                    import json
                    changes = []
                    if new_title != old_title:
                        changes.append({
                            "field": "ประเภทงาน",
                            "old": old_title or "(ไม่ระบุ)",
                            "new": new_title or "(ไม่ระบุ)"
                        })
                    if new_notes != old_notes:
                        changes.append({
                            "field": "หมายเหตุ",
                            "old": old_notes or "(ไม่มี)",
                            "new": new_notes or "(ไม่มี)"
                        })
                    note_text = json.dumps(changes, ensure_ascii=False) if changes else "แก้ไขข้อมูลคำขอ"
                    try:
                        ServiceRequestLog.objects.create(
                            request=sr,
                            action="แก้ไขคำขอ",
                            actor=request.user,
                            note=note_text,
                        )
                    except Exception:
                        pass
                    messages.success(request, "บันทึกข้อมูลเรียบร้อยแล้ว")
                else:
                    messages.info(request, "ไม่มีการเปลี่ยนแปลงที่จะบันทึก")

            return redirect("service_request_detail", sr_id=sr.id)

        # GET: prepare template context
        attachments = sr.attachments.all()
        logs = sr.logs.order_by("-created_at")

        # Use persisted preferred_service_date field (do not parse from notes)
        try:
            desired_service_date = getattr(sr, 'preferred_service_date', None)
        except Exception:
            desired_service_date = None

        # Technicians for inline WO creation
        try:
            from .models import Technician
            technicians = Technician.objects.filter(active=True).order_by("name")
        except Exception:
            technicians = []

        return render(
            request,
            "workorder/service_request_detail.html",
            {
                "sr": sr,
                "attachments": attachments,
                "logs": logs,
                "desired_service_date": desired_service_date,
                "technicians": technicians,
            },
        )
    




@login_required
def api_daily_summary(request):
    """API endpoint to get daily work summary matching technician_availability_report logic.
    
    Query parameters:
    - date: Date in YYYY-MM-DD format
    
    Returns JSON with summary counts for work/OT/leave/shift/holiday matching report calculations.
    """
    from datetime import datetime
    
    from .models import ShiftSchedule, Technician, TechnicianAvailability
    
    try:
        date_str = request.GET.get("date")
        
        if not date_str:
            return JsonResponse({"error": "date parameter required"}, status=400)
        
        # Parse the date
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            return JsonResponse({"error": "invalid date format"}, status=400)
        
        # Load holidays from localStorage (if available from request header/cookie)
        # For now we'll check if the date is in a known holiday set
        # This matches the report's holidaySet logic
        holidays = []
        try:
            import json
            # Try to get holidays from a session or cache if available
            # For simplicity, we'll skip this for now and let frontend handle
            pass
        except Exception:
            pass
        holiday_set = set(h.get('date') for h in holidays if h.get('date'))
        
        # Helper functions matching report logic
        def classify_status_full(avail):
            status = (avail.status or '').lower()
            note = (avail.note or '').lower()
            text = f"{status} {note}"
            
            if any(keyword in text for keyword in ['off_', 'holiday', 'วันหยุด', 'หยุดราชการ', 'หยุด']):
                return 'holiday'
            if any(keyword in text for keyword in ['leave', 'ลา', 'ลากิจ', 'ป่วย', 'sick', 'personal', 'vacation', 'annual', 'พักร้อน']):
                return 'leave'
            return 'work'
        
        def get_ot_hours_robust(avail):
            raw_note = (avail.note or '')
            raw_status = (avail.status or '')
            note = raw_note.lower()
            status = raw_status.lower()
            
            import re
            # Try various patterns to extract OT hours
            patterns = [
                r'\bot[:_\s-]*?(\d{1,2})\b',
                r'\bovertime[:_\s-]*?(\d{1,2})\b',
                r'(\d{1,2})\s*(h|hr|hrs|hour|hours|ชม|ชั่วโมง)\b',
                r'\((\d{1,2})\)',
            ]
            
            for pattern in patterns:
                m = re.search(pattern, note) or re.search(pattern, status)
                if m:
                    return int(m.group(1))
            
            # If contains 'ot' or 'overtime' but no number, default to 4
            if any(kw in note or kw in status for kw in ['ot', 'overtime', 'ล่วงเวลา', 'ล่วง']):
                return 4
            
            return 0
        
        # Get all active technicians
        technicians = Technician.objects.filter(active=True)
        
        # Get availability records for this date
        availabilities = TechnicianAvailability.objects.filter(
            date=target_date
        ).select_related('technician')
        
        # Get shift schedules for this date
        shifts = ShiftSchedule.objects.filter(
            date=target_date
        ).select_related('technician')
        
        # Calculate stats matching report logic
        stats = {
            'totalWork': 0,
            'totalOT': 0,
            'totalOTWeekday': 0,
            'totalOTHoliday': 0,
            'totalOTHours': 0,
            'totalOTWeekdayHours': 0,
            'totalOTHolidayHours': 0,
            'totalWorkHours': 0,
            'totalLeave': 0,
            'totalLeaveSick': 0,
            'totalLeavePersonal': 0,
            'totalLeaveAnnual': 0,
            'totalLeaveOther': 0,
            'totalHoliday': 0,
            'totalShift': 0,
            'totalShiftMorning': 0,
            'totalShiftAfternoon': 0,
            'totalShiftHours': 0,
            'totalShiftMorningHours': 0,
            'totalShiftAfternoonHours': 0,
        }
        
        # Process availability records
        for avail in availabilities:
            status_cat = classify_status_full(avail)
            ot_hours = get_ot_hours_robust(avail)
            
            if status_cat == 'leave':
                stats['totalLeave'] += 1
                note_lower = (avail.note or '').lower()
                status_lower = (avail.status or '').lower()
                text = f"{note_lower} {status_lower}"
                
                if 'ป่วย' in text or 'sick' in text:
                    stats['totalLeaveSick'] += 1
                elif 'ลากิจ' in text or 'personal' in text:
                    stats['totalLeavePersonal'] += 1
                elif 'พักร้อน' in text or 'annual' in text or 'vacation' in text:
                    stats['totalLeaveAnnual'] += 1
                else:
                    stats['totalLeaveOther'] += 1
            else:
                if ot_hours > 0:
                    # Count OT occurrences and OT hours
                    stats['totalOT'] += 1
                    stats['totalOTHours'] += ot_hours

                    # Determine if weekday or holiday OT
                    is_holiday_ot = date_str in holiday_set or target_date.weekday() >= 5

                    if is_holiday_ot:
                        # OT on holiday/weekend: count as OT only
                        stats['totalOTHoliday'] += 1
                        stats['totalOTHolidayHours'] += ot_hours
                    else:
                        # OT on a regular weekday: count OT plus the regular work hours
                        stats['totalOTWeekday'] += 1
                        stats['totalOTWeekdayHours'] += ot_hours

                        # Also count the base work for that technician on a weekday
                        # so totalWorkHours reflects normal 7h work plus the OT hours.
                        stats['totalWork'] += 1
                        stats['totalWorkHours'] += 7
                else:
                    if status_cat == 'work':
                        stats['totalWork'] += 1
                        stats['totalWorkHours'] += 7
                    elif status_cat == 'holiday':
                        stats['totalHoliday'] += 1
        
        # Process shift schedules
        for shift in shifts:
            stats['totalShift'] += 1
            if shift.shift == 'morning':
                stats['totalShiftMorning'] += 1
                stats['totalShiftMorningHours'] += 8
                stats['totalShiftHours'] += 8
            elif shift.shift == 'afternoon':
                stats['totalShiftAfternoon'] += 1
                stats['totalShiftAfternoonHours'] += 16
                stats['totalShiftHours'] += 16
            else:
                stats['totalShiftHours'] += 8
        
        # Add metadata
        stats['date'] = date_str
        stats['totalTechnicians'] = technicians.count()
        stats['totalRecords'] = availabilities.count()
        
        return JsonResponse(stats, safe=False)
    
    except Exception as e:
        logger.error(f"Error in api_daily_summary: {str(e)}", exc_info=True)
        return JsonResponse({"error": str(e)}, status=500)


@login_required
def daily_capacity_management(request):
    """View for daily capacity management page."""
    return render(request, "maintenance/daily_capacity.html")


@login_required
@login_required
def capacity_summary(request):
    """Admin read-only summary: for the next N days, show per-skill available capacity."""
    from datetime import timedelta

    from django.utils import timezone

    from .models import (Holiday, MaintenanceAppointment, MaintenanceCapacity,
                         Technician, TechnicianAvailability)

    days = int(request.GET.get("days", 14))
    days = max(1, min(days, 60))
    today = timezone.localtime().date()
    end_date = today + timedelta(days=days - 1)

    # ── Active technicians + skill map ────────────────────────────────────────
    all_techs = list(Technician.objects.filter(active=True).prefetch_related("skills_m"))
    skill_to_techs: dict[str, list] = {}
    for tech in all_techs:
        for skill in tech.skills_m.all():
            skill_to_techs.setdefault(skill.label, []).append(tech)
    skill_labels = sorted(skill_to_techs.keys())

    # Headers for template: [{label, tech_count}]
    skill_headers = [
        {"label": lbl, "tech_count": len(skill_to_techs[lbl])}
        for lbl in skill_labels
    ]

    # ── Holidays: {date: name} ───────────────────────────────────────────────
    holiday_map: dict = {
        h.date: h.name
        for h in Holiday.objects.filter(is_active=True, date__gte=today, date__lte=end_date)
    }

    # ── TechnicianAvailability: {date: {tech_id: status}} ───────────────────
    avail_map: dict = {}
    for av in TechnicianAvailability.objects.filter(
        date__gte=today, date__lte=end_date
    ).values("technician_id", "date", "status"):
        avail_map.setdefault(av["date"], {})[av["technician_id"]] = av["status"]

    # ── Appointments in range (one row per appointment) ──────────────────────
    raw_appts = list(
        MaintenanceAppointment.objects.filter(
            scheduled_date__gte=today, scheduled_date__lte=end_date
        ).values("scheduled_date", "equipment__equipment_name_EN",
                 "equipment__equipment_type")
    )
    # Build: {date: {skill_label: count}}  (substring match both ways)
    appt_by_skill: dict = {}
    for a in raw_appts:
        d = a["scheduled_date"]
        eq_name = (a["equipment__equipment_name_EN"] or "").lower()
        eq_type = (a["equipment__equipment_type"] or "").lower()
        for lbl in skill_labels:
            lbl_l = lbl.lower()
            if lbl_l in eq_name or eq_name in lbl_l or lbl_l in eq_type or eq_type in lbl_l:
                appt_by_skill.setdefault(d, {}).setdefault(lbl, 0)
                appt_by_skill[d][lbl] += 1

    # ── Configured capacity override: {(date, label): capacity} ─────────────
    cap_map: dict = {
        (c["date"], c["equipment_type"]): c["capacity"]
        for c in MaintenanceCapacity.objects.filter(
            date__gte=today, date__lte=end_date
        ).values("date", "equipment_type", "capacity")
    }

    # ── Thai weekday names ────────────────────────────────────────────────────
    WEEKDAY_TH = ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"]

    # Status values that mean the technician IS available for work.
    # The DB uses 'work_regular' / 'work_shift' / 'work_ot' (legacy 'working'/'overtime').
    WORKING_STATUSES = {"working", "work_regular", "overtime", "work_shift", "work_ot"}

    # ── Build grid: ALL days (working + non-working) ─────────────────────────
    date_rows = []
    for i in range(days):
        d = today + timedelta(days=i)
        is_holiday = d in holiday_map
        is_weekend = d.weekday() >= 5
        is_working = not is_holiday and not is_weekend

        day_avail = avail_map.get(d, {})
        # If the date has ANY availability record, only explicitly-working techs count.
        # If no records at all for this date, assume everyone is working (planning mode).
        date_has_records = d in avail_map
        cols = []
        day_total_cap = 0
        day_total_booked = 0

        for lbl in skill_labels:
            if not is_working:
                cols.append({"eq_type": lbl, "techs": 0, "capacity": 0,
                             "booked": 0, "remaining": 0, "is_working": False,
                             "tech_total": len(skill_to_techs.get(lbl, [])),
                             "off_count": 0})
                continue

            tech_list = skill_to_techs.get(lbl, [])
            tech_total = len(tech_list)
            avail_count = 0
            avail_cap = 0
            off_count = 0
            tech_names = []
            off_names = []
            for tech in tech_list:
                if date_has_records:
                    # Records exist → unknown tech = not working
                    status = day_avail.get(tech.id, "off")
                else:
                    # No records for this date → planning mode, assume working
                    status = day_avail.get(tech.id, "working")
                if status in WORKING_STATUSES:
                    avail_count += 1
                    avail_cap += tech.per_day_capacity
                    tech_names.append(tech.name)
                else:
                    off_count += 1
                    off_names.append(tech.name)

            configured = cap_map.get((d, lbl))
            effective = min(configured, avail_cap) if configured is not None else avail_cap
            booked = appt_by_skill.get(d, {}).get(lbl, 0)
            remaining = max(0, effective - booked)

            day_total_cap += effective
            day_total_booked += booked
            cols.append({
                "eq_type": lbl,
                "techs": avail_count,
                "tech_total": tech_total,
                "off_count": off_count,
                "tech_names": tech_names,
                "off_names": off_names,
                "capacity": effective,
                "booked": booked,
                "remaining": remaining,
                "is_working": True,
            })

        date_rows.append({
            "date": d,
            "weekday_th": WEEKDAY_TH[d.weekday()],
            "cols": cols,
            "is_working": is_working,
            "is_weekend": is_weekend,
            "is_holiday": is_holiday,
            "holiday_name": holiday_map.get(d, ""),
            "total_cap": day_total_cap,
            "total_booked": day_total_booked,
            "total_remaining": max(0, day_total_cap - day_total_booked),
        })

    context = {
        "date_rows": date_rows,
        "skill_headers": skill_headers,
        "days": days,
        "today": today,
        "total_active_techs": len(all_techs),
        "has_availability_records": bool(avail_map),
        "working_days": sum(1 for r in date_rows if r["is_working"]),
    }
    return render(request, "maintenance/capacity_summary.html", context)


def api_available_dates_for_equipment(request):
    """API endpoint to find available appointment dates based on equipment type and technician skills.
    
    This is the SMART API that connects:
    1. Equipment → equipment_type
    2. equipment_type → Technicians with matching skills (skills_m)
    3. Technicians → TechnicianAvailability (who is working/available)
    4. Date capacity → MaintenanceCapacity + existing MaintenanceAppointment count
    
    Query parameters:
    - equipment_id: Equipment ID to get type from (optional if equipment_type provided)
    - equipment_type: Direct equipment type string (optional if equipment_id provided)
    - start_date: Start date for search (default: today)
    - days: Number of days to search (default: 21)
    
    Returns JSON with available dates and capacity info.
    """
    from datetime import datetime, timedelta
    
    from django.db.models import Count, Q
    from django.utils import timezone
    
    from .models import (Equipment_list, Holiday, MaintenanceAppointment,
                         MaintenanceCapacity, Technician, TechnicianAvailability)
    
    try:
        # Get parameters
        equipment_id = request.GET.get("equipment_id", "").strip()
        equipment_type = request.GET.get("equipment_type", "").strip()
        start_date_str = request.GET.get("start_date", "")
        days = int(request.GET.get("days", 21))
        
        # Limit days to prevent abuse
        days = min(days, 60)
        
        # Determine equipment info for matching
        equipment_name_en = ""
        equipment_code = ""
        if equipment_id and not equipment_type:
            # Build query - only include id lookup if equipment_id is numeric
            q_filter = Q(equipment_id=equipment_id)
            if equipment_id.isdigit():
                q_filter |= Q(id=int(equipment_id))
            eq = Equipment_list.objects.filter(q_filter).first()
            if eq:
                equipment_type = eq.equipment_type or ""
                equipment_name_en = eq.equipment_name_EN or ""
                equipment_code = eq.equipment_code or ""
        
        # Build search terms for technician skill matching
        # Priority: equipment_name_EN > equipment_code > equipment_type
        search_terms = []
        if equipment_name_en:
            search_terms.append(equipment_name_en)
        if equipment_code:
            search_terms.append(equipment_code)
        if equipment_type:
            search_terms.append(equipment_type)
        
        if not search_terms:
            return JsonResponse({
                "success": False,
                "error": "กรุณาระบุ equipment_id หรือ equipment_type",
                "dates": []
            })
        
        # Parse start date
        if start_date_str:
            try:
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d").date()
            except ValueError:
                start_date = timezone.localtime().date()
        else:
            start_date = timezone.localtime().date()
        
        # Get holidays set
        holiday_dates = set(
            Holiday.objects.filter(
                is_active=True,
                date__gte=start_date,
                date__lte=start_date + timedelta(days=days)
            ).values_list("date", flat=True)
        )
        
        # Find technicians with matching skills
        # Try each search term until we find matching technicians
        matching_techs = None
        matched_term = None
        
        for term in search_terms:
            techs = Technician.objects.filter(
                active=True
            ).filter(
                Q(skills_m__label__icontains=term) |
                Q(skills__icontains=term)
            ).distinct()
            if techs.exists():
                matching_techs = techs
                matched_term = term
                break
        
        if not matching_techs:
            # No match found with any term
            return JsonResponse({
                "success": True,
                "equipment_type": equipment_type,
                "equipment_name_EN": equipment_name_en,
                "equipment_code": equipment_code,
                "search_terms_tried": search_terms,
                "message": f"ไม่พบช่างที่มีทักษะตรงกับเครื่องมือนี้",
                "dates": [],
                "technicians_count": 0
            })
        
        tech_ids = list(matching_techs.values_list("id", flat=True))
        tech_capacity_map = {t.id: t.per_day_capacity for t in matching_techs}
        
        # Get availability records for these technicians in date range
        end_date = start_date + timedelta(days=days)
        availabilities = TechnicianAvailability.objects.filter(
            technician_id__in=tech_ids,
            date__gte=start_date,
            date__lte=end_date
        ).values("technician_id", "date", "status")
        
        # Build availability map: {date: {tech_id: status}}
        avail_map = {}
        for av in availabilities:
            d = av["date"]
            if d not in avail_map:
                avail_map[d] = {}
            avail_map[d][av["technician_id"]] = av["status"]
        
        # Get existing appointments count per date
        appointments = MaintenanceAppointment.objects.filter(
            scheduled_date__gte=start_date,
            scheduled_date__lte=end_date,
            equipment__equipment_type=equipment_type
        ).values("scheduled_date").annotate(count=Count("id"))
        appt_count_map = {a["scheduled_date"]: a["count"] for a in appointments}
        
        # Get configured capacity per date
        capacities = MaintenanceCapacity.objects.filter(
            date__gte=start_date,
            date__lte=end_date,
            equipment_type=equipment_type
        ).values("date", "capacity")
        capacity_map = {c["date"]: c["capacity"] for c in capacities}
        
        # Build result for each day
        available_dates = []
        
        for i in range(days):
            current_date = start_date + timedelta(days=i)
            
            # Skip weekends (Saturday=5, Sunday=6)
            if current_date.weekday() >= 5:
                continue
            
            # Skip holidays
            if current_date in holiday_dates:
                continue
            
            # Count available technicians for this date
            day_avail = avail_map.get(current_date, {})
            date_has_records = current_date in avail_map
            # DB uses work_regular/work_shift/work_ot; also support legacy working/overtime
            working_statuses = {"working", "work_regular", "overtime", "work_shift", "work_ot"}
            available_tech_count = 0
            available_tech_capacity = 0
            
            for tech_id in tech_ids:
                if date_has_records:
                    status = day_avail.get(tech_id, "off")  # records exist → unknown = off
                else:
                    status = day_avail.get(tech_id, "working")  # no records → planning mode
                if status in working_statuses:
                    available_tech_count += 1
                    available_tech_capacity += tech_capacity_map.get(tech_id, 1)
            
            if available_tech_count == 0:
                continue
            
            # Calculate remaining capacity
            configured_capacity = capacity_map.get(current_date)
            existing_appts = appt_count_map.get(current_date, 0)
            
            if configured_capacity is not None:
                # Use minimum of configured capacity and tech capacity
                effective_capacity = min(configured_capacity, available_tech_capacity)
            else:
                # No configured limit, use tech capacity
                effective_capacity = available_tech_capacity
            
            remaining = max(0, effective_capacity - existing_appts)
            
            if remaining > 0:
                available_dates.append({
                    "date": current_date.isoformat(),
                    "weekday": ["จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์", "อาทิตย์"][current_date.weekday()],
                    "available_technicians": available_tech_count,
                    "total_capacity": effective_capacity,
                    "booked": existing_appts,
                    "remaining": remaining,
                    "is_configured": configured_capacity is not None
                })
        
        return JsonResponse({
            "success": True,
            "equipment_type": equipment_type,
            "equipment_name_EN": equipment_name_en,
            "equipment_code": equipment_code,
            "matched_skill_term": matched_term,
            "technicians_with_skill": len(tech_ids),
            "search_range": {
                "start": start_date.isoformat(),
                "end": end_date.isoformat(),
                "days": days
            },
            "dates": available_dates,
            "total_available_dates": len(available_dates)
        })
    
    except Exception as e:
        logger.error(f"Error in api_available_dates_for_equipment: {str(e)}", exc_info=True)
        return JsonResponse({"success": False, "error": str(e), "dates": []}, status=500)


@permission_required("cmms.change_workorder", raise_exception=True)
def delete_workorder_attachment(request, wo_id, att_id):
    """Delete a WorkOrderAttachment (trash icon). POST only."""
    if request.method != "POST":
        from django.http import HttpResponseNotAllowed
        return HttpResponseNotAllowed(["POST"])
    from cmms.models import WorkOrderAttachment
    import os
    try:
        att = WorkOrderAttachment.objects.get(id=att_id, workorder_id=wo_id)
        # Delete the physical file as well
        if att.file and att.file.name:
            try:
                file_path = att.file.path
                att.delete()
                if os.path.isfile(file_path):
                    os.remove(file_path)
            except Exception:
                att.delete()
        else:
            att.delete()
        messages.success(request, "ลบเอกสารแนบแล้ว")
    except WorkOrderAttachment.DoesNotExist:
        messages.error(request, "ไม่พบเอกสารแนบที่ต้องการลบ")
    return redirect("workorder_detail", wo_id=wo_id)


# ===================== Checklist บำรุงรักษา/สอบเทียบ =====================

def _equipment_identity_options():
    """รายชื่อชนิดเครื่องมือสำหรับ checkbox — ดึงจากข้อมูลกลางหน้า Master Data
    (/master/category/equipments) เพื่อให้ตรงกับรายการที่แอดมินดูแลไว้ที่เดียว"""
    from .models import MasterItem

    values = set()
    for v in MasterItem.objects.filter(category="equipments", active=True).values_list("label", flat=True):
        v = (v or "").strip()
        if v:
            values.add(v)
    return sorted(values)


def _equipment_identity_candidates(equipment):
    """ค่าที่ใช้เทียบหาแบบฟอร์ม Checklist ที่ตรงกับอุปกรณ์ชิ้นนี้ เรียงจากเจาะจงที่สุด
    (equipment_name_EN/TH มักระบุชนิดเครื่องมือได้ชัดกว่า equipment_type ซึ่งมักเป็นหมวดกว้าง)"""
    if not equipment:
        return []
    candidates = []
    for val in (equipment.equipment_name_EN, equipment.equipment_name_TH, equipment.equipment_type):
        val = (val or "").strip()
        if val and val not in candidates:
            candidates.append(val)
    return candidates


def _can_manage_checklist_templates(user):
    return user.is_authenticated and (
        user.is_staff or user.is_superuser or user.has_perm("cmms.change_workorder")
    )


def _resolve_workorder_type_codes(wo):
    """คืนชุด code ประเภทงาน (เช่น {'calibration'}) — ใบงานจริงมักเก็บประเภทผ่าน
    workorder_types (M2M ไป MasterItem) ไม่ใช่ผ่าน legacy field workorder_type อย่างเดียว
    จึงต้องไล่ตรวจแบบเดียวกับ workorderTypePills เพื่อให้ตรงกับที่ผู้ใช้เห็นจริงบนหน้าจอ"""
    from .models import MasterItem, WorkOrder as _WorkOrder
    from django.db.models import Q

    codes = set()
    try:
        wo_types = list(wo.workorder_types.all())
    except Exception:
        wo_types = []
    label_to_code = {label: code for code, label in _WorkOrder.WORKORDER_TYPE_CHOICES}
    for t in wo_types:
        if t.code:
            codes.add(t.code)
        elif t.label in label_to_code:
            codes.add(label_to_code[t.label])

    if not codes and wo.workorder_type and wo.workorder_type != "other":
        codes.add(wo.workorder_type)

    if not codes:
        try:
            from cmms.views_reports import _TITLE_TYPE_Q as _GLOBAL_TITLE_Q
        except Exception:
            _GLOBAL_TITLE_Q = {}
        for m in MasterItem.objects.filter(category="workorder_type", active=True):
            label = (m.label or "").strip()
            q_expr = None
            if m.code and m.code in _GLOBAL_TITLE_Q:
                q_expr = _GLOBAL_TITLE_Q[m.code]
                if label:
                    q_expr = q_expr | Q(title__icontains=label)
            elif label:
                q_expr = Q(title__icontains=label)
            if q_expr and m.code and _WorkOrder.objects.filter(id=wo.id).filter(q_expr).exists():
                codes.add(m.code)
    return codes


@user_passes_test(_can_manage_checklist_templates)
def checklist_templates(request):
    """รายการแบบฟอร์ม Checklist ทั้งหมด พร้อมกรองตามประเภทใบงาน/ประเภทเครื่องมือ"""
    from .models import ChecklistTemplate, WorkOrder as _WorkOrder

    qs = ChecklistTemplate.objects.all().order_by(
        "workorder_type", "equipment_type", "name"
    )
    filter_wo_type = request.GET.get("workorder_type", "").strip()
    filter_eq_type = request.GET.get("equipment_type", "").strip()
    if filter_wo_type:
        qs = qs.filter(workorder_type=filter_wo_type)
    if filter_eq_type:
        qs = qs.filter(equipment_type__icontains=filter_eq_type)

    return render(
        request,
        "checklist/checklist_templates.html",
        {
            "templates": qs,
            "workorder_type_choices": _WorkOrder.WORKORDER_TYPE_CHOICES,
            "filter_workorder_type": filter_wo_type,
            "filter_equipment_type": filter_eq_type,
        },
    )


@user_passes_test(_can_manage_checklist_templates)
def checklist_template_form(request, pk=None):
    """สร้าง/แก้ไขแบบฟอร์ม Checklist (ชื่อ + รายการตรวจสอบ) — รายการถูกส่งมาเป็น JSON เดียวคล้ายรูปแบบ spare_parts_data"""
    from .models import ChecklistTemplate, ChecklistTemplateItem, WorkOrder as _WorkOrder
    from django.db import transaction
    from django.db.models.deletion import ProtectedError
    import json as _json

    template = get_object_or_404(ChecklistTemplate, pk=pk) if pk else None
    workorder_type_choices = _WorkOrder.WORKORDER_TYPE_CHOICES
    result_type_choices = ChecklistTemplateItem.RESULT_TYPE_CHOICES
    equipment_type_options = _equipment_identity_options()
    # เคยถูกใช้กรอกผลแล้วหรือยัง — ถ้าใช่ การบันทึกแก้ไขจะสร้างเวอร์ชันใหม่แทนการเขียนทับ
    template_in_use = bool(template and template.workorder_checklists.exists())

    if request.method == "POST":
        name = request.POST.get("name", "").strip()
        selected_equipment_types = [v.strip() for v in request.POST.getlist("equipment_types") if v.strip()]
        equipment_type = ",".join(selected_equipment_types)
        workorder_type = request.POST.get("workorder_type", "").strip()
        description = request.POST.get("description", "").strip()
        active = request.POST.get("active") == "on"
        items_json = request.POST.get("items_data", "[]")

        errors = []
        if not name:
            errors.append("กรุณาระบุชื่อแบบฟอร์ม")
        valid_wo_types = {c[0] for c in workorder_type_choices}
        if workorder_type not in valid_wo_types:
            errors.append("กรุณาเลือกประเภทใบงาน")

        try:
            raw_items = _json.loads(items_json)
            if not isinstance(raw_items, list):
                raw_items = []
        except (ValueError, TypeError):
            raw_items = []
            errors.append("ข้อมูลรายการ checklist ไม่ถูกต้อง")

        def _to_decimal(v):
            try:
                return round(float(v), 4) if v not in (None, "") else None
            except (TypeError, ValueError):
                return None

        valid_result_types = {c[0] for c in result_type_choices}
        cleaned_items = []
        for idx, raw in enumerate(raw_items):
            if not isinstance(raw, dict):
                continue
            item_name = str(raw.get("name", "")).strip()
            if not item_name:
                continue
            result_type = raw.get("result_type") or "check"
            if result_type not in valid_result_types:
                result_type = "check"
            cleaned_items.append(
                {
                    "name": item_name,
                    "description": str(raw.get("description", "")).strip(),
                    "result_type": result_type,
                    "unit": str(raw.get("unit", "")).strip(),
                    "standard_value": _to_decimal(raw.get("standard_value")),
                    "tolerance": _to_decimal(raw.get("tolerance")),
                    "is_required": bool(raw.get("is_required", True)),
                    "order": idx,
                }
            )

        if not cleaned_items:
            errors.append("กรุณาเพิ่มอย่างน้อย 1 รายการตรวจสอบ")

        form_data = {
            "name": name,
            "equipment_type": equipment_type,
            "selected_equipment_types": selected_equipment_types,
            "workorder_type": workorder_type,
            "description": description,
            "active": active,
        }

        if errors:
            for e in errors:
                messages.error(request, e)
            return render(
                request,
                "checklist/checklist_template_form.html",
                {
                    "template": template,
                    "workorder_type_choices": workorder_type_choices,
                    "result_type_choices": result_type_choices,
                    "form_data": form_data,
                    "items_for_json": cleaned_items or raw_items,
                    "equipment_type_options": equipment_type_options,
                    "template_in_use": template_in_use,
                },
            )

        revised = False
        try:
            with transaction.atomic():
                if template is None:
                    template = ChecklistTemplate.objects.create(
                        name=name,
                        equipment_type=equipment_type,
                        workorder_type=workorder_type,
                        description=description,
                        active=active,
                        created_by=request.user,
                    )
                    for item in cleaned_items:
                        ChecklistTemplateItem.objects.create(template=template, **item)
                elif template.workorder_checklists.exists():
                    # เคยถูกใช้กรอกผลในใบงานแล้ว แก้ไขรายการเดิมไม่ได้ (ผลเก่าอ้างอิงรายการ/เกณฑ์ชุดเดิม)
                    # จึงสร้างเป็นเวอร์ชันใหม่แทน และเก็บเวอร์ชันเดิมไว้เป็นประวัติ
                    old_template = template
                    new_template = ChecklistTemplate.objects.create(
                        name=name,
                        equipment_type=equipment_type,
                        workorder_type=workorder_type,
                        description=description,
                        active=active,
                        created_by=request.user,
                        version=old_template.version + 1,
                    )
                    for item in cleaned_items:
                        ChecklistTemplateItem.objects.create(template=new_template, **item)
                    old_template.replaced_by = new_template
                    old_template.active = False
                    old_template.save(update_fields=["replaced_by", "active"])
                    template = new_template
                    revised = True
                else:
                    template.name = name
                    template.equipment_type = equipment_type
                    template.workorder_type = workorder_type
                    template.description = description
                    template.active = active
                    template.save()
                    template.items.all().delete()
                    for item in cleaned_items:
                        ChecklistTemplateItem.objects.create(template=template, **item)
        except ProtectedError:
            messages.error(
                request,
                "แบบฟอร์มนี้เคยถูกใช้กรอกผลในใบงานแล้ว ไม่สามารถแก้ไขรายการเดิมได้ "
                "กรุณาสร้างแบบฟอร์มใหม่ หรือปิดการใช้งาน (ปุ่มลบ) แบบฟอร์มนี้แทน",
            )
            return render(
                request,
                "checklist/checklist_template_form.html",
                {
                    "template": template,
                    "workorder_type_choices": workorder_type_choices,
                    "result_type_choices": result_type_choices,
                    "form_data": form_data,
                    "items_for_json": cleaned_items,
                    "equipment_type_options": equipment_type_options,
                    "template_in_use": template_in_use,
                },
            )

        if revised:
            messages.success(
                request,
                f"แบบฟอร์มนี้เคยถูกใช้กรอกผลแล้ว จึงบันทึกเป็นเวอร์ชันใหม่ (v{template.version}) แทน "
                "และเก็บเวอร์ชันเดิมไว้เป็นประวัติ (ปิดการใช้งานอัตโนมัติ)",
            )
        else:
            messages.success(request, "บันทึกแบบฟอร์ม Checklist เรียบร้อยแล้ว")
        return redirect("checklist_templates")

    items_for_json = []
    form_data = {
        "name": "",
        "equipment_type": "",
        "selected_equipment_types": [],
        "workorder_type": "",
        "description": "",
        "active": True,
    }
    if template:
        form_data = {
            "name": template.name,
            "equipment_type": template.equipment_type,
            "selected_equipment_types": template.equipment_type_list,
            "workorder_type": template.workorder_type,
            "description": template.description,
            "active": template.active,
        }
        items_for_json = [
            {
                "name": it.name,
                "description": it.description,
                "result_type": it.result_type,
                "unit": it.unit,
                "standard_value": str(it.standard_value) if it.standard_value is not None else "",
                "tolerance": str(it.tolerance) if it.tolerance is not None else "",
                "is_required": it.is_required,
            }
            for it in template.items.all()
        ]

    return render(
        request,
        "checklist/checklist_template_form.html",
        {
            "template": template,
            "workorder_type_choices": workorder_type_choices,
            "result_type_choices": result_type_choices,
            "form_data": form_data,
            "items_for_json": items_for_json,
            "equipment_type_options": equipment_type_options,
            "template_in_use": template_in_use,
        },
    )


@user_passes_test(_can_manage_checklist_templates)
@require_http_methods(["POST"])
def checklist_template_delete(request, pk):
    """ลบแบบฟอร์ม — ถ้าเคยถูกใช้งานแล้วจะปิดการใช้งาน (active=False) แทนการลบจริง"""
    from .models import ChecklistTemplate
    from django.db.models.deletion import ProtectedError

    template = get_object_or_404(ChecklistTemplate, pk=pk)
    try:
        template.delete()
        messages.success(request, "ลบแบบฟอร์มเรียบร้อยแล้ว")
    except ProtectedError:
        template.active = False
        template.save(update_fields=["active"])
        messages.warning(request, "แบบฟอร์มนี้ถูกใช้งานในใบงานแล้ว จึงปิดการใช้งานแทนการลบ")
    return redirect("checklist_templates")


@login_required
def workorder_checklist_fill(request, wo_id):
    """หน้ากรอกผล Checklist ของใบงานหนึ่งใบ — ผูกอัตโนมัติกับแบบฟอร์มตามประเภทใบงาน+ประเภทเครื่องมือ"""
    from .models import ChecklistTemplate, WorkOrder, WorkOrderChecklist, WorkOrderChecklistResult

    wo = get_object_or_404(WorkOrder, pk=wo_id)

    can_fill = request.user == wo.assigned_to or request.user.has_perm("cmms.change_workorder")
    if not can_fill:
        messages.error(request, "คุณไม่มีสิทธิ์กรอก Checklist ของใบงานนี้")
        return redirect("workorder_detail", wo_id=wo.id)

    checklist = getattr(wo, "checklist", None)
    if checklist is None:
        equipment_candidates = _equipment_identity_candidates(wo.equipment) or [""]
        type_codes = _resolve_workorder_type_codes(wo) or {wo.workorder_type}
        template = None
        for code in type_codes:
            for equipment_type in equipment_candidates:
                template = ChecklistTemplate.find_for(code, equipment_type)
                if template:
                    break
            if template:
                break
        if template is None:
            messages.info(
                request,
                "ยังไม่มีแบบฟอร์ม Checklist สำหรับใบงานประเภทนี้/เครื่องมือประเภทนี้ "
                "กรุณาสร้างแบบฟอร์มก่อนที่เมนู 'แบบฟอร์ม Checklist'",
            )
            return redirect("workorder_detail", wo_id=wo.id)
        checklist = WorkOrderChecklist.objects.create(workorder=wo, template=template)

    items = list(checklist.template.items.all())

    if request.method == "POST":
        errors = []
        to_save = []
        for item in items:
            prefix = f"item_{item.id}_"
            result_choice = request.POST.get(prefix + "result", "").strip()
            measured_raw = request.POST.get(prefix + "measured", "").strip()
            note = request.POST.get(prefix + "note", "").strip()

            measured_value = None
            if item.result_type == "measurement" and measured_raw:
                try:
                    measured_value = round(float(measured_raw), 4)
                except ValueError:
                    errors.append(f"{item.name}: ค่าที่วัดได้ต้องเป็นตัวเลข")
                    continue

            if item.is_required:
                filled = (
                    (item.result_type == "check" and result_choice)
                    or (item.result_type == "measurement" and measured_value is not None)
                    or (item.result_type == "text" and note)
                )
                if not filled:
                    errors.append(f"{item.name}: จำเป็นต้องกรอก")
                    continue

            is_pass = None
            if item.result_type == "measurement":
                is_pass = item.evaluate(measured_value)
            elif item.result_type == "check":
                if result_choice == "pass":
                    is_pass = True
                elif result_choice == "fail":
                    is_pass = False

            to_save.append(
                {
                    "item": item,
                    "result_choice": result_choice if item.result_type == "check" else "",
                    "measured_value": measured_value,
                    "is_pass": is_pass,
                    "note": note,
                }
            )

        if errors:
            for e in errors:
                messages.error(request, e)
        else:
            for data in to_save:
                WorkOrderChecklistResult.objects.update_or_create(
                    checklist=checklist,
                    template_item=data["item"],
                    defaults={
                        "result_choice": data["result_choice"],
                        "measured_value": data["measured_value"],
                        "is_pass": data["is_pass"],
                        "note": data["note"],
                    },
                )
            if request.POST.get("finalize") == "1":
                checklist.completed_by = request.user
                checklist.completed_at = timezone.now()
                checklist.save(update_fields=["completed_by", "completed_at"])
                messages.success(request, "บันทึกและยืนยันผล Checklist เรียบร้อยแล้ว")
            else:
                messages.success(request, "บันทึกผล Checklist เรียบร้อยแล้ว (ฉบับร่าง)")
            return redirect("workorder_checklist_fill", wo_id=wo.id)

    results_by_item = {r.template_item_id: r for r in checklist.results.all()}
    rows = [{"item": item, "result": results_by_item.get(item.id)} for item in items]

    return render(
        request,
        "checklist/workorder_checklist_fill.html",
        {
            "wo": wo,
            "checklist": checklist,
            "rows": rows,
            "can_finalize": can_fill,
        },
    )

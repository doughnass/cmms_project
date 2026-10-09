import os
import re
import uuid

from django.conf import settings
from django.db import models
from django.utils import timezone

# Create your models here.


# Extended profile for additional user attributes
class Profile(models.Model):
    APPROVAL_PENDING = "pending"
    APPROVAL_APPROVED = "approved"
    APPROVAL_REJECTED = "rejected"
    APPROVAL_STATUS_CHOICES = [
        (APPROVAL_PENDING, "รอตรวจสอบ"),
        (APPROVAL_APPROVED, "อนุมัติแล้ว"),
        (APPROVAL_REJECTED, "ปฏิเสธ"),
    ]

    user = models.OneToOneField(
        settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name="profile"
    )
    telephone = models.CharField(max_length=50, blank=True, default="")
    department = models.CharField(max_length=200, blank=True, default="")
    position = models.CharField(max_length=200, blank=True, default="")
    user_type = models.CharField(max_length=100, blank=True, default="")

    # Account approval workflow. Self-registered accounts start as "pending";
    # accounts created by staff/admin tooling default to "approved" so they
    # are not accidentally blocked from logging in.
    approval_status = models.CharField(
        max_length=20, choices=APPROVAL_STATUS_CHOICES, default=APPROVAL_APPROVED
    )
    approved_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="approved_profiles",
    )
    approved_at = models.DateTimeField(null=True, blank=True)
    rejection_reason = models.TextField(blank=True, default="")

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = "Profile"
        verbose_name_plural = "Profiles"
        permissions = [
            ("can_approve_accounts", "Can review and approve new account registrations"),
        ]

    def __str__(self):
        return f"Profile for {self.user}"


# Ensure a Profile exists for each User
from django.db.models.signals import post_save
from django.dispatch import receiver


@receiver(post_save, sender=settings.AUTH_USER_MODEL)
def create_user_profile(sender, instance, created, **kwargs):
    if created:
        Profile.objects.create(user=instance)


@receiver(post_save, sender=settings.AUTH_USER_MODEL)
def save_user_profile(sender, instance, **kwargs):
    try:
        instance.profile.save()
    except Profile.DoesNotExist:
        # profile might not exist in some edge cases; create it
        Profile.objects.get_or_create(user=instance)


# ทะเบียนเครื่องมือแพทย์
def equipment_image_upload_path(instance, filename):
    extension = os.path.splitext(filename)[1].lower()
    equipment_id = re.sub(r"[^A-Za-z0-9_-]+", "_", instance.equipment_id).strip("_")
    return f"equipment/{equipment_id or instance.pk}{extension}"


class Equipment_list(models.Model):
    # ข้อมูลทั่วไป
    equipment_id = models.CharField(
        max_length=20, unique=True, db_index=True
    )  # MEDCMU-000000
    equipment_code = models.CharField(max_length=20, blank=True, default="")  # EML
    # Make these fields optional at the model level so the admin add form
    # and other programmatic creates can omit non-essential values. The
    # front-end `equipment_form.html` already enforces the visible
    # required fields; the DB should be permissive for admin/scripted
    # operations to avoid blocking creation.
    equipment_name_EN = models.CharField(max_length=100, blank=True, default="")  # ชื่อ(ไทย)
    equipment_name_TH = models.CharField(max_length=100, blank=True, default="")  # ชื่อ(English)
    equipment_brand = models.CharField(max_length=100, blank=True, default="")  # ยี่ห้อ
    equipment_model = models.CharField(max_length=100, blank=True, default="")  # รุ่น
    equipment_sn = models.CharField(max_length=100, blank=True, default="")  # serial no.
    equipment_gov = models.CharField(max_length=100, blank=True, default="")  # เลขครุภัณฑ์
    equipment_price = models.IntegerField(null=True, blank=True)  # ราคา
    equipment_photo = models.CharField(max_length=100, blank=True, default="")  # รูป
    equipment_image = models.FileField(
        upload_to=equipment_image_upload_path, blank=True
    )
    equipment_type = models.CharField(max_length=100, blank=True, default="")  # ประเภท
    equipment_life = models.IntegerField(null=True, blank=True)  # อายุการใช้งาน

    # การรับประกัน
    equipment_waranty_date = models.DateField(null=True, blank=True)
    equipment_waranty_due = models.DateField(null=True, blank=True)
    equipment_distributor_name = models.CharField(max_length=100, blank=True, default="")  # ชื่อบริษัทผู้จำหน่าย/ผู้รับประกัน
    equipment_distributor_tel = models.CharField(max_length=50, blank=True, default="")  # เบอร์ติดต่อ บริษัทผู้จำหน่าย/ผู้รับประกัน
    # ผู้ดูแล (เดิมมีฟิลด์รายละเอียดระดับหน่วยงาน แต่ระบบปัจจุบันใช้เฉพาะ
    # equipment_owner_customer / equipment_user_customer จาก master data)
    # ฟิลด์ระดับย่อยของผู้ดูแลถูกลบเพื่อให้ schema ตรงกับฟอร์มหน้าเว็บ
    # แผนการบำรุงรักษา
    equipment_pm_fq = models.IntegerField(null=True, blank=True)  # ความถี่การ PM
    equipment_pm_due = models.DateField(null=True, blank=True)  # วันที่ครบกำหนด PM
    # Estimated maintenance time in hours for this equipment (optional).
    # If not provided, a default per-task-type value will be used.
    estimated_hours = models.DecimalField(
        null=True,
        blank=True,
        max_digits=5,
        decimal_places=2,
        help_text='ป้อนจำนวนชั่วโมงที่คาดว่าใช้ต่อชิ้น (เช่น 1.5)'
    )
    # แผนสอบเทียบ
    equipment_cal_fq = models.IntegerField(null=True, blank=True)  # ความถี่การ CAL
    equipment_cal_due = models.DateField(null=True, blank=True)  # วันที่ครบกำหนด CAL
    # flags to indicate whether this equipment requires PM and/or CAL
    requires_pm = models.BooleanField(default=True)
    requires_cal = models.BooleanField(default=False)
    # ข้อมูลเจ้าของ
    # ปัจจุบันฟอร์มใช้เพียง `equipment_owner_customer` (Master data label)
    equipment_owner_customer = models.CharField(max_length=100, blank=True, default="")  # เจ้าของ องค์กร คณะแพทย์/คณะวิทย์
    # ข้อมูลผู้ใช้งาน
    # ฟอร์มปัจจุบันใช้เพียง `equipment_user_customer` (Master data label)
    equipment_user_customer = models.CharField(max_length=100, blank=True, default="")  # ผู้ใช้งาน คณะแพทย์/คณะวิทย์/รพ.ต่างๆ
    # ผู้ดูแล/ผู้ให้บริการ
    equipment_service_provider = models.CharField(max_length=100, blank=True, default="")  # ผู้ดูแล/ผู้ให้บริการ
    # การขึ้นทะเบียน
    equipment_register_username = models.CharField(max_length=100, blank=True, default="")  # ผู้แจ้งขึ้นทะเบียน
    equipment_register_adminname = models.CharField(max_length=100, blank=True, default="")  # ชื่อแอดมินผู้ขึ้นทะเบียน
    equipment_register_date = models.DateField(auto_now_add=True)  # วันที่ขึ้นทะเบียน
    # อื่นๆ
    equipment_note = models.CharField(max_length=300, blank=True, default="")
    
    # ประวัติการขึ้นทะเบียนและแก้ไข (Audit Trail)
    created_at = models.DateTimeField(auto_now_add=True, verbose_name='สร้างเมื่อ')
    created_by = models.CharField(max_length=100, blank=True, default='', verbose_name='สร้างโดย')
    updated_at = models.DateTimeField(auto_now=True, verbose_name='แก้ไขล่าสุดเมื่อ')
    updated_by = models.CharField(max_length=100, blank=True, default='', verbose_name='แก้ไขล่าสุดโดย')

    def __str__(self):
        # Be defensive: some fields were removed/are optional. Use getattr
        # with defaults to avoid AttributeError for older rows or schema
        # mismatches until migrations are applied.
        eid = getattr(self, 'equipment_id', '') or ''
        name = getattr(self, 'equipment_name_EN', '') or ''
        owner = getattr(self, 'equipment_owner_customer', '') or ''
        user = getattr(self, 'equipment_user_customer', '') or ''
        parts = [p for p in (eid, name, owner, user) if p]
        return ", ".join(parts) if parts else eid or name or super().__str__()


def equipment_profile_file_upload_path(instance, filename):
    extension = os.path.splitext(os.path.basename(filename))[1].lower()
    equipment_id = re.sub(
        r"[^A-Za-z0-9_-]+", "_", instance.equipment.equipment_id
    ).strip("_")
    return (
        f"equipment_profile/{equipment_id or instance.equipment_id}/"
        f"{instance.section}/{uuid.uuid4().hex}{extension}"
    )


class EquipmentProfileFile(models.Model):
    IMAGE = "image"
    DOCUMENT = "document"
    SECTION_CHOICES = [
        (IMAGE, "รูปภาพ"),
        (DOCUMENT, "เอกสาร/คู่มือ"),
    ]

    equipment = models.ForeignKey(
        Equipment_list, on_delete=models.CASCADE, related_name="profile_files"
    )
    section = models.CharField(max_length=12, choices=SECTION_CHOICES)
    file = models.FileField(upload_to=equipment_profile_file_upload_path)
    display_name = models.CharField(max_length=255)
    uploaded_by = models.CharField(max_length=255, blank=True, default="")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["-uploaded_at", "-id"]

    @property
    def is_image(self):
        return os.path.splitext(self.file.name)[1].lower() in {
            ".jpg",
            ".jpeg",
            ".png",
            ".gif",
            ".webp",
            ".bmp",
        }

    def __str__(self):
        return f"{self.equipment.equipment_id}: {self.display_name}"


# ประวัติการขึ้นทะเบียนและแก้ไขอุปกรณ์ (Equipment Registration and Edit History)
class EquipmentHistory(models.Model):
    ACTION_TYPES = [
        ('CREATE', 'สร้างทะเบียน'),
        ('UPDATE', 'แก้ไขทะเบียน'),
        ('DELETE', 'ลบทะเบียน'),
        ('FILE_UPLOAD', 'อัปโหลดไฟล์'),
        ('FILE_DELETE', 'ลบไฟล์'),
    ]
    
    equipment_id = models.CharField(max_length=20, db_index=True, verbose_name='รหัสเครื่อง')
    action_type = models.CharField(max_length=20, choices=ACTION_TYPES, verbose_name='ประเภทการดำเนินการ')
    action_by = models.CharField(max_length=100, verbose_name='ดำเนินการโดย')
    action_at = models.DateTimeField(auto_now_add=True, verbose_name='เวลาที่ดำเนินการ')
    ip_address = models.GenericIPAddressField(null=True, blank=True, verbose_name='IP Address')
    
    # JSON fields to store what changed
    changed_fields = models.JSONField(null=True, blank=True, verbose_name='ฟิลด์ที่เปลี่ยนแปลง')  # List of field names
    old_values = models.JSONField(null=True, blank=True, verbose_name='ค่าเดิม')  # Dict of old values
    new_values = models.JSONField(null=True, blank=True, verbose_name='ค่าใหม่')  # Dict of new values
    
    notes = models.TextField(blank=True, default='', verbose_name='หมายเหตุ')
    
    class Meta:
        verbose_name = 'ประวัติการขึ้นทะเบียนอุปกรณ์'
        verbose_name_plural = 'ประวัติการขึ้นทะเบียนอุปกรณ์'
        ordering = ['-action_at']
        indexes = [
            models.Index(fields=['equipment_id', '-action_at'], name='equipment_history_idx'),
            models.Index(fields=['action_type'], name='action_type_idx'),
        ]
    
    def __str__(self):
        return f"{self.equipment_id} - {self.get_action_type_display()} โดย {self.action_by} ({self.action_at.strftime('%Y-%m-%d %H:%M')})"


# ลูกค้า
class Customer_list(models.Model):
    customer_id = models.CharField(max_length=20)  # CUST-000000
    customer_units = models.CharField(max_length=100)  # หน่วยงาน
    customer_section = models.CharField(max_length=100)  # สังกัด
    customer_department = models.CharField(max_length=100)  # ฝ่าย/งาน
    customer_semi_department = models.CharField(
        max_length=100, blank=True, default=""
    )  # งาน (ในฝ่าย)
    customer_division = models.CharField(max_length=100)  # สวนดอก/ศรีพัฒน์/CMEx/คณะต่างๆ
    customer_customer = models.CharField(max_length=100)  # องค์กร คณะแพทย์/คณะวิทย์
    customer_name = models.CharField(max_length=100)  # ชื่อผู้ติดต่อ
    customer_tel = models.CharField(max_length=20)  # เบอร์ติดต่อ
    customer_email = models.CharField(max_length=100)  # อีเมล
    customer_address = models.CharField(max_length=300)  # ที่อยู่

    def __str__(self):
        return (
            self.customer_id
            + ", "
            + self.customer_units
            + ", "
            + self.customer_section
            + ", "
            + self.customer_semi_department
            + ", "
            + self.customer_department
            + ", "
            + self.customer_division
            + ", "
            + self.customer_customer
            + ", "
            + self.customer_name
            + ", "
            + self.customer_tel
            + ", "
            + self.customer_email
            + ", "
            + self.customer_address
        )


# Work orders / service requests
class WorkOrder(models.Model):
    PRIORITY_CHOICES = [
        ("low", "Low"),
        ("medium", "Medium"),
        ("high", "High"),
    ]
    WORKORDER_TYPE_CHOICES = [
        ("repair", "ซ่อม"),
        ("maintenance", "บำรุงรักษา"),
        ("calibration", "สอบเทียบ"),
        ("installation", "ติดตั้ง"),
        ("consultation", "ให้คำปรึกษา"),
        ("other", "อื่น ๆ"),
    ]
    STATUS_CHOICES = [
        ("open", "Open"),
        ("assigned", "Assigned"),
        ("in_progress", "In Progress"),
        ("on_hold", "On Hold"),
        ("completed", "Completed"),
        ("verified", "Verified"),
        ("closed", "Closed"),
    ]

    HOLD_REASON_CHOICES = [
        ("waiting_parts", "รออะไหล่"),
        ("waiting_outsource", "รอผู้ให้บริการภายนอก"),
        ("other", "อื่นๆ"),
    ]

    title = models.CharField(max_length=200)
    workorder_type = models.CharField(
        max_length=50, choices=WORKORDER_TYPE_CHOICES, default="other"
    )
    description = models.TextField(blank=True)
    # allow multiple types (e.g. maintenance + calibration together)
    from django.conf import settings as _settings

    workorder_types = models.ManyToManyField(
        "MasterItem", blank=True, related_name="workorders"
    )
    # whether the user accepted the system's suggested title at creation time
    accepted_title_suggestion = models.BooleanField(default=False)
    equipment = models.ForeignKey(
        Equipment_list, null=True, blank=True, on_delete=models.SET_NULL
    )
    reported_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="reported_workorders",
    )
    reported_at = models.DateTimeField(auto_now_add=True)
    assigned_to = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="assigned_workorders",
    )
    # officer who reviewed/accepted the request
    accepted_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="accepted_workorders",
    )
    accepted_at = models.DateTimeField(null=True, blank=True)
    accepted_note = models.TextField(blank=True, default="")
    priority = models.CharField(
        max_length=10, choices=PRIORITY_CHOICES, default="medium"
    )
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="open")
    planned_start = models.DateTimeField(null=True, blank=True)
    planned_end = models.DateTimeField(null=True, blank=True)
    actual_start = models.DateTimeField(null=True, blank=True)
    actual_end = models.DateTimeField(null=True, blank=True)
    notes = models.TextField(blank=True)
    # technician report / notes when work completed
    technician_report = models.TextField(blank=True, default="")
    # engineer verification
    verified_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="verified_workorders",
    )
    verified_at = models.DateTimeField(null=True, blank=True)
    verified_note = models.TextField(blank=True, default="")
    # On Hold tracking
    hold_reason = models.CharField(
        max_length=30, choices=HOLD_REASON_CHOICES, blank=True, default="",
        verbose_name="เหตุผลการระงับ",
    )
    hold_note = models.TextField(blank=True, default="", verbose_name="หมายเหตุการระงับ")
    # Follow-up tracking (เครื่องใช้งานได้ แต่มีอะไหล่ค้าง)
    follow_up_needed = models.BooleanField(
        default=False, verbose_name="ต้องติดตาม",
        help_text="ใบงานเสร็จแล้ว แต่มีอะไหล่รอจัดซื้อ — เมื่ออะไหล่มาจะสร้างใบงานติดตาม",
    )
    follow_up_for = models.ForeignKey(
        "self", null=True, blank=True, on_delete=models.SET_NULL,
        related_name="follow_up_orders",
        verbose_name="ใบงานต้นเรื่อง",
        help_text="ลิงก์กลับไปหาใบงานต้นเรื่องที่มีอะไหล่ค้าง",
    )
    # Outsource tracking
    outsource_vendor = models.CharField(
        max_length=200, blank=True, default="",
        verbose_name="ผู้ให้บริการภายนอก",
    )
    outsource_cost = models.DecimalField(
        max_digits=12, decimal_places=2, null=True, blank=True,
        verbose_name="ค่าใช้จ่าย Outsource",
    )
    outsource_expected_date = models.DateField(
        null=True, blank=True,
        verbose_name="วันที่คาดว่า Outsource เสร็จ",
    )
    updated_at = models.DateTimeField(auto_now=True)

    def get_priority_label(self):
        """Return a human-friendly label for `priority`.

        Maps legacy values used by the SR -> WO flow and falls back to
        the model's `get_priority_display()` for normal choices.
        """
        mapping = {
            "intime": "ในเวลา",
            "overtime": "ล่วงเวลา",
        }
        if self.priority in mapping:
            return mapping[self.priority]
        try:
            return self.get_priority_display()
        except Exception:
            return self.priority or ""

    class Meta:
        ordering = ["-reported_at"]

    def __str__(self):
        return f"WO#{self.id} {self.title} [{self.status}]"


class WorkOrderLog(models.Model):
    """Audit log for workorder workflow actions."""

    workorder = models.ForeignKey(
        WorkOrder, on_delete=models.CASCADE, related_name="logs"
    )
    action = models.CharField(
        max_length=100
    )  # e.g., 'accepted', 'assigned', 'started', 'completed', 'verified', 'closed'
    actor = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    note = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["created_at"]

    def __str__(self):
        return (
            f"WO#{self.workorder.id} {self.action} by {self.actor} at {self.created_at}"
        )


class WorkOrderComment(models.Model):
    workorder = models.ForeignKey(
        WorkOrder, on_delete=models.CASCADE, related_name="comments"
    )
    author = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    text = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["created_at"]

    def __str__(self):
        return f"Comment by {self.author} on WO#{self.workorder.id}"


class WorkOrderAttachment(models.Model):
    DOCUMENT_TYPE_CHOICES = [
        ("pm_report", "รายงานผลการบำรุงรักษา"),
        ("cal_report", "รายงานผลการสอบเทียบ"),
        ("repair_report", "รายงานผลการซ่อม"),
        ("other", "อื่นๆ"),
    ]

    workorder = models.ForeignKey(
        WorkOrder, on_delete=models.CASCADE, related_name="attachments"
    )
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    document_name = models.CharField(max_length=255, blank=True, default="")
    document_type = models.CharField(
        max_length=20,
        choices=DOCUMENT_TYPE_CHOICES,
        default="other",
    )
    file = models.FileField(upload_to="workorder_attachments/")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        label = self.document_name or self.file.name
        return f"Attachment {label} for WO#{self.workorder.id}"


# Reusable master data model
class MasterItem(models.Model):
    """Generic key/value for master data.

    category: grouping name (e.g. 'equipment_type', 'department', 'manufacturer')
    code: optional short code
    label: human-friendly label
    description: optional text description
    meta: flexible JSON field for per-category attributes (defined by MasterFieldDefinition)
    """

    category = models.CharField(max_length=100, db_index=True)
    code = models.CharField(max_length=50, blank=True, default="")
    label = models.CharField(max_length=200)
    description = models.TextField(blank=True, default="")
    active = models.BooleanField(default=True)
    order = models.IntegerField(default=0)
    parent = models.ForeignKey(
        "self", null=True, blank=True, on_delete=models.CASCADE, related_name="children"
    )
    meta = models.JSONField(null=True, blank=True, verbose_name="ข้อมูลเพิ่มเติม (Meta)")

    class Meta:
        ordering = ["category", "order", "label"]
        unique_together = (("category", "code"),)

    def __str__(self):
        # Prefer showing the human label (and code if present). Avoid showing category prefix
        if self.code:
            return f"{self.label} ({self.code})"
        return f"{self.label}"


class MasterCategory(models.Model):
    """Defines schema/metadata for each MasterItem category.

    Allows admins to configure which extra fields (beyond label/code/description)
    each category should have, stored in MasterItem.meta.
    """

    key = models.CharField(
        max_length=100, unique=True, db_index=True, verbose_name="Category Key"
    )
    label = models.CharField(max_length=200, verbose_name="ชื่อหมวดหมู่")
    description = models.TextField(blank=True, default="", verbose_name="คำอธิบาย")
    # Optional template for how to display items of this category as select options
    # e.g. "{{ label }}{% if code %} ({{ code }}){% endif %}{% if meta.floor %} — ชั้น {{ meta.floor }}{% endif %}"
    display_template = models.TextField(
        blank=True,
        default="",
        verbose_name="Display Label Template",
        help_text="Django template snippet (context: label, code, meta). Leave blank for default.",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["label"]
        verbose_name = "Master Category"
        verbose_name_plural = "Master Categories"

    def __str__(self):
        return f"{self.label} ({self.key})"


class MasterFieldDefinition(models.Model):
    """Defines a custom field for items in a specific MasterCategory.

    These fields are stored in MasterItem.meta as key-value pairs.
    """

    FIELD_TYPE_CHOICES = [
        ("string", "String (text)"),
        ("text", "Text (textarea)"),
        ("integer", "Integer"),
        ("decimal", "Decimal"),
        ("boolean", "Boolean (checkbox)"),
        ("date", "Date"),
        ("select", "Select (dropdown)"),
    ]

    category = models.ForeignKey(
        MasterCategory,
        on_delete=models.CASCADE,
        related_name="fields",
        verbose_name="หมวดหมู่",
    )
    name = models.CharField(
        max_length=100,
        verbose_name="Field Name (machine key)",
        help_text='Machine-safe name (lowercase, underscores, no spaces). e.g. "floor_number"',
    )
    label = models.CharField(max_length=200, verbose_name="Label (แสดงใน UI)")
    field_type = models.CharField(
        max_length=20,
        choices=FIELD_TYPE_CHOICES,
        default="string",
        verbose_name="Field Type",
    )
    required = models.BooleanField(default=False, verbose_name="Required (จำเป็น)")
    # For select type: list of choices. Format: [{"value":"a","label":"A"}, ...] or simple ["A","B","C"]
    choices = models.JSONField(
        null=True,
        blank=True,
        verbose_name="Choices (for select type)",
        help_text='JSON array of choices. e.g. ["A","B","C"] or [{"value":"a","label":"A"}]',
    )
    default_value = models.CharField(
        max_length=200, blank=True, default="", verbose_name="Default Value"
    )
    order = models.IntegerField(default=0, verbose_name="Order (ลำดับ)")
    help_text = models.CharField(
        max_length=400, blank=True, default="", verbose_name="Help Text"
    )
    # Whether to show this field value in list/option previews
    visible_in_preview = models.BooleanField(
        default=False, verbose_name="Visible in Preview"
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["category", "order", "label"]
        unique_together = (("category", "name"),)
        verbose_name = "Master Field Definition"
        verbose_name_plural = "Master Field Definitions"

    def __str__(self):
        return f"{self.category.label}.{self.name} ({self.get_field_type_display()})"


# Service requests submitted by customers (separate from internal WorkOrder)
class ServiceRequest(models.Model):
    REQUEST_TYPE_CHOICES = [
        ("other", "ทั่วไป"),
        ("maintenance", "ซ่อมบำรุง"),
        ("calibration", "ปรับเทียบ"),
        ("inspection", "ตรวจสอบ"),
        ("installation", "ติดตั้ง"),
    ]
    STATUS_CHOICES = [
        ("new", "New"),
        ("reviewed", "Reviewed by officer"),
        ("converted", "Converted to WorkOrder"),
        ("rejected", "Rejected"),
        ("deleted", "Deleted"),
        ("closed", "Closed"),
    ]

    title = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    equipment = models.ForeignKey(
        Equipment_list, null=True, blank=True, on_delete=models.SET_NULL
    )
    # Preferred service date (explicit DB field; do not store in notes)
    preferred_service_date = models.DateField(null=True, blank=True)
    # optional customer info if the submitter is anonymous
    customer_name = models.CharField(max_length=200, blank=True, default="")
    customer_email = models.CharField(max_length=200, blank=True, default="")
    customer_position = models.CharField(max_length=200, blank=True, default="")
    customer_telephone = models.CharField(max_length=50, blank=True, default="")
    customer_department = models.CharField(max_length=200, blank=True, default="")
    requested_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="service_requests",
    )
    requested_at = models.DateTimeField(auto_now_add=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="new")
    # request type / ประเภทคำขอ
    request_type = models.CharField(
        max_length=50, choices=REQUEST_TYPE_CHOICES, default="other"
    )
    notes = models.TextField(blank=True, default="")
    converted_to = models.OneToOneField(
        "WorkOrder",
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="origin_request",
    )

    @property
    def display_id(self):
        """Human-friendly Service Request ID in format: <BuddhistYear>-<5-digit>

        Uses the SR's `requested_at` year when available, otherwise falls back
        to the current year. The numeric part is the zero-padded primary key.
        """
        try:
            year = self.requested_at.year if self.requested_at else timezone.now().year
        except Exception:
            year = timezone.now().year
        buddhist_year = year + 543 - 2500  # Convert to Buddhist year and subtract 2500 for short format
        num = int(self.id) if self.id else 0
        return f"SR-{buddhist_year}-{num:05d}"

    class Meta:
        ordering = ["-requested_at"]

    def __str__(self):
        return f"SR#{self.id} {self.title} [{self.status}]"


class ServiceRequestLog(models.Model):
    request = models.ForeignKey(
        ServiceRequest, on_delete=models.CASCADE, related_name="logs"
    )
    action = models.CharField(max_length=100)
    actor = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    note = models.TextField(blank=True, default="")
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["created_at"]

    def __str__(self):
        return (
            f"SR#{self.request.id} {self.action} by {self.actor} at {self.created_at}"
        )


class ServiceRequestAttachment(models.Model):
    request = models.ForeignKey(
        ServiceRequest, on_delete=models.CASCADE, related_name="attachments"
    )
    uploaded_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    file = models.FileField(upload_to="service_request_attachments/")
    uploaded_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"SR Attachment {self.file.name} for SR#{self.request.id}"


# Scheduling models for planned maintenance
class MaintenanceCapacity(models.Model):
    """Defines how many devices of a given equipment_type can be serviced on a particular date.

    e.g. on 2025-10-20, technicians available can service up to 5 'ventilator' devices and 3 'xray' devices.
    """

    date = models.DateField(db_index=True)
    equipment_type = models.CharField(
        max_length=100,
        help_text="Matches Equipment_list.equipment_type or MasterItem label/code",
    )
    capacity = models.IntegerField(default=0)

    class Meta:
        unique_together = (("date", "equipment_type"),)
        ordering = ["date", "equipment_type"]

    def __str__(self):
        return f"{self.date} – {self.equipment_type}: {self.capacity}"


class MaintenanceAppointment(models.Model):
    """Represents a scheduled maintenance appointment for an equipment item."""

    equipment = models.ForeignKey(
        Equipment_list,
        on_delete=models.CASCADE,
        related_name="maintenance_appointments",
    )
    scheduled_date = models.DateField(db_index=True)
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL
    )
    created_at = models.DateTimeField(auto_now_add=True)
    notes = models.TextField(blank=True, default="")
    assigned_technician = models.ForeignKey(
        "Technician",
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="appointments",
        verbose_name="ช่างที่ได้รับมอบหมาย",
    )

    class Meta:
        ordering = ["scheduled_date"]

    def __str__(self):
        return f"Appointment {self.equipment} on {self.scheduled_date}"


# Technicians and availability ---------------------------------------------
class Technician(models.Model):
    """Represents a technician (staff) who can perform maintenance tasks.

    We keep a simple model with name, user (optional link to auth user), skills (free-text / comma-separated
    master item codes), and active flag.
    """

    PREFIX_CHOICES = [
        ("", "ไม่ระบุ"),
        ("นาย", "นาย"),
        ("นาง", "นาง"),
        ("นางสาว", "นางสาว"),
        ("other", "อื่นๆ (ระบุเอง)"),
    ]

    prefix_name = models.CharField(
        max_length=50,
        choices=PREFIX_CHOICES,
        blank=True,
        default="",
        verbose_name="คำนำหน้า",
    )
    prefix_name_custom = models.CharField(
        max_length=50,
        blank=True,
        default="",
        verbose_name="คำนำหน้า (ระบุเอง)",
        help_text='ใช้เมื่อเลือก "อื่นๆ (ระบุเอง)"',
    )
    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="tech_profile",
    )
    name = models.CharField(max_length=200)
    # simple skill tag: can match Equipment_list.equipment_type or MasterItem.code/label
    skills = models.CharField(max_length=200, blank=True, default="")
    phone = models.CharField(max_length=50, blank=True, default="")
    # how many devices this technician can handle per day (default 1)
    per_day_capacity = models.IntegerField(default=1)
    # สามารถขึ้นเวรได้หรือไม่
    can_work_shifts = models.BooleanField(default=False, verbose_name="สามารถขึ้นเวรได้")
    active = models.BooleanField(default=True)
    # new ManyToMany relation to MasterItem for normalized skills (category='technician_skill')
    skills_m = models.ManyToManyField(
        "MasterItem",
        blank=True,
        related_name="technicians",
        limit_choices_to={"category": "equipments"},
        verbose_name="ทักษะ / ประเภทอุปกรณ์",
    )

    def __str__(self):
        return self.name


class TechnicianAvailability(models.Model):
    """Availability record for a technician on a specific date.

    status: working / off / holiday / overtime
    optional note for reason (e.g., sick, training)
    """

    STATUS_CHOICES = [
        ("working", "Working"),
        ("off", "Off"),
        ("holiday", "Holiday"),
        ("overtime", "Overtime"),
    ]
    technician = models.ForeignKey(
        Technician, on_delete=models.CASCADE, related_name="availabilities"
    )
    date = models.DateField(db_index=True)
    status = models.CharField(max_length=20, choices=STATUS_CHOICES, default="working")
    note = models.CharField(max_length=300, blank=True, default="")

    class Meta:
        unique_together = (("technician", "date"),)
        ordering = ["date", "technician_id"]

    def __str__(self):
        return f"{self.technician} on {self.date}: {self.status}"


class Holiday(models.Model):
    """
    Model for storing public holidays
    """

    date = models.DateField(unique=True, db_index=True, verbose_name="วันที่")
    name = models.CharField(max_length=200, verbose_name="ชื่อวันหยุด")
    description = models.TextField(blank=True, default="", verbose_name="รายละเอียด")
    is_active = models.BooleanField(default=True, verbose_name="เปิดใช้งาน")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="สร้างเมื่อ")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="แก้ไขเมื่อ")
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="created_holidays",
        verbose_name="สร้างโดย",
    )

    class Meta:
        ordering = ["date"]
        verbose_name = "วันหยุดราชการ"
        verbose_name_plural = "วันหยุดราชการ"
        indexes = [
            models.Index(fields=["date", "is_active"]),
        ]

    def __str__(self):
        return f"{self.date.strftime('%d/%m/%Y')} - {self.name}"

    @property
    def year(self):
        return self.date.year

    @property
    def month(self):
        return self.date.month

    @property
    def thai_date(self):
        """Return date in Thai Buddhist calendar"""
        return self.date.strftime("%d/%m/") + str(self.date.year + 543)


class ShiftSchedule(models.Model):
    """
    Model for storing shift schedules (เวรเช้า/เวรบ่าย)
    """

    SHIFT_CHOICES = [
        ("morning", "เวรเช้า"),
        ("afternoon", "เวรบ่าย"),
    ]

    technician = models.ForeignKey(
        Technician,
        on_delete=models.CASCADE,
        related_name="shift_schedules",
        verbose_name="ช่าง",
    )
    date = models.DateField(db_index=True, verbose_name="วันที่")
    shift = models.CharField(max_length=20, choices=SHIFT_CHOICES, verbose_name="เวร")
    note = models.CharField(
        max_length=300, blank=True, default="", verbose_name="หมายเหตุ"
    )
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="สร้างเมื่อ")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="แก้ไขเมื่อ")
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="created_shift_schedules",
        verbose_name="สร้างโดย",
    )

    class Meta:
        unique_together = (("date", "shift"),)  # แต่ละวัน แต่ละเวร มีช่างได้คนเดียว
        ordering = ["date", "shift"]
        verbose_name = "ตารางเวร"
        verbose_name_plural = "ตารางเวร"
        indexes = [
            models.Index(fields=["date", "shift"]),
            models.Index(fields=["technician", "date"]),
        ]

    def __str__(self):
        return f"{self.date} - {self.get_shift_display()}: {self.technician.name}"


# Default durations for different maintenance task types
class MaintenanceTaskDuration(models.Model):
    TASK_CHOICES = [
        ("maintenance", "Maintenance / PM"),
        ("calibration", "Calibration"),
        ("repair", "Repair"),
    ]

    task_type = models.CharField(max_length=30, choices=TASK_CHOICES, unique=True)
    hours = models.DecimalField(max_digits=5, decimal_places=2, default=1)

    class Meta:
        verbose_name = "Default maintenance task duration"
        verbose_name_plural = "Default maintenance task durations"

    def __str__(self):
        return f"{self.get_task_type_display()}: {self.hours} ชม."


# ──────────────────────────────────────────────────────────────────────────────
# ทะเบียนอะไหล่คงคลัง (Spare Parts Inventory)
# ──────────────────────────────────────────────────────────────────────────────
class SparePart(models.Model):
    """อะไหล่สำหรับเครื่องมือแพทย์ — ใช้จัดการคลังอะไหล่"""

    code = models.CharField(
        max_length=50, unique=True, db_index=True, verbose_name="รหัสอะไหล่"
    )
    name = models.CharField(max_length=200, verbose_name="ชื่ออะไหล่")
    description = models.TextField(blank=True, default="", verbose_name="รายละเอียด")
    category = models.ForeignKey(
        MasterItem,
        null=True,
        blank=True,
        on_delete=models.SET_NULL,
        related_name="spare_parts",
        verbose_name="หมวดหมู่",
        help_text="เลือกจาก Master Data (หมวดหมู่อะไหล่)",
    )
    unit = models.CharField(
        max_length=50, blank=True, default="ชิ้น", verbose_name="หน่วยนับ"
    )
    quantity = models.IntegerField(default=0, verbose_name="จำนวนคงเหลือ")
    min_quantity = models.IntegerField(
        default=0, verbose_name="จำนวนขั้นต่ำ",
        help_text="แจ้งเตือนเมื่อจำนวนคงเหลือต่ำกว่านี้",
    )
    unit_cost = models.DecimalField(
        max_digits=12, decimal_places=2, null=True, blank=True,
        verbose_name="ราคาต่อหน่วย (บาท)",
    )
    location = models.CharField(
        max_length=200, blank=True, default="", verbose_name="ตำแหน่งจัดเก็บ"
    )
    # ผู้จำหน่าย / supplier
    supplier = models.CharField(
        max_length=200, blank=True, default="", verbose_name="ผู้จำหน่าย"
    )
    # อุปกรณ์ที่ใช้ร่วมกับอะไหล่นี้ (optional)
    compatible_equipment = models.ManyToManyField(
        Equipment_list,
        blank=True,
        related_name="spare_parts",
        verbose_name="อุปกรณ์ที่ใช้ร่วมได้",
    )
    active = models.BooleanField(default=True, verbose_name="ใช้งาน")
    notes = models.TextField(blank=True, default="", verbose_name="หมายเหตุ")
    # JSON store for multi-select equipment associations (names, brands, models)
    equipment_meta = models.JSONField(
        default=dict, blank=True,
        verbose_name="ข้อมูลเครื่องมือที่ใช้งาน",
        help_text='{"equipment_names":[],"equipment_brands":[],"equipment_models":[]}',
    )

    created_at = models.DateTimeField(auto_now_add=True, verbose_name="สร้างเมื่อ")
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="created_spare_parts",
        verbose_name="สร้างโดย",
    )
    updated_at = models.DateTimeField(auto_now=True, verbose_name="แก้ไขล่าสุด")
    updated_by = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        null=True, blank=True,
        on_delete=models.SET_NULL,
        related_name="updated_spare_parts",
        verbose_name="แก้ไขโดย",
    )

    class Meta:
        ordering = ["code"]
        verbose_name = "อะไหล่"
        verbose_name_plural = "อะไหล่"

    def __str__(self):
        return f"{self.code} — {self.name}"

    @property
    def is_low_stock(self):
        return self.quantity <= self.min_quantity


class WorkOrderSparePart(models.Model):
    """อะไหล่ที่เบิกใช้ในใบงาน — รองรับทั้งอะไหล่จากคลังและอะไหล่อื่นๆ"""

    workorder = models.ForeignKey(
        WorkOrder, on_delete=models.CASCADE, related_name="used_spare_parts"
    )
    spare_part = models.ForeignKey(
        SparePart, on_delete=models.CASCADE, related_name="workorder_usages",
        null=True, blank=True,
        help_text="เลือกจากคลัง หรือเว้นว่างสำหรับอะไหล่อื่นๆ",
    )
    # ฟิลด์สำหรับอะไหล่อื่นๆ (ไม่มีในคลัง)
    custom_name = models.CharField(
        max_length=200, blank=True, default="",
        verbose_name="ชื่ออะไหล่ (อื่นๆ)",
    )
    custom_unit_cost = models.DecimalField(
        max_digits=12, decimal_places=2, null=True, blank=True,
        verbose_name="ราคาต่อหน่วย (อื่นๆ)",
    )
    custom_unit = models.CharField(
        max_length=50, blank=True, default="ชิ้น",
        verbose_name="หน่วยนับ (อื่นๆ)",
    )
    quantity = models.PositiveIntegerField(default=1, verbose_name="จำนวนที่เบิก")
    is_out_of_stock = models.BooleanField(
        default=False, verbose_name="รอจัดซื้อ",
        help_text="อะไหล่ไม่มีในคลัง ต้องทำเรื่องจัดซื้อ",
    )
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        ordering = ["id"]
        verbose_name = "อะไหล่ที่ใช้ในใบงาน"
        verbose_name_plural = "อะไหล่ที่ใช้ในใบงาน"

    @property
    def display_name(self):
        if self.spare_part:
            return self.spare_part.name
        return self.custom_name or "อะไหล่อื่นๆ"

    @property
    def display_code(self):
        if self.spare_part:
            return self.spare_part.code
        return "อื่นๆ"

    @property
    def display_unit(self):
        if self.spare_part:
            return self.spare_part.unit
        return self.custom_unit or "ชิ้น"

    @property
    def display_unit_cost(self):
        if self.spare_part:
            return self.spare_part.unit_cost
        return self.custom_unit_cost

    @property
    def is_custom(self):
        return self.spare_part is None

    def __str__(self):
        name = self.spare_part.code if self.spare_part else self.custom_name
        return f"WO#{self.workorder_id} — {name} x{self.quantity}"


# ===================== Checklist บำรุงรักษา/สอบเทียบ (แทนไฟล์ Excel/PDF แนบ) =====================

class ChecklistTemplate(models.Model):
    """แบบฟอร์ม checklist ที่แก้ไขได้ ผูกกับประเภทเครื่องมือ + ประเภทใบงาน (PM/สอบเทียบ/ฯลฯ)."""

    name = models.CharField(max_length=200, verbose_name="ชื่อแบบฟอร์ม")
    # เก็บได้หลายชนิด คั่นด้วย ',' (เลือกจาก checkbox ไม่ให้พิมพ์เองเพื่อกันชื่อไม่ตรงกัน)
    # เว้นว่าง = ใช้ได้กับเครื่องมือทุกประเภท (แบบฟอร์มกลาง)
    equipment_type = models.CharField(
        max_length=500, blank=True, default="",
        verbose_name="ประเภทเครื่องมือ",
        help_text="เว้นว่างเพื่อใช้เป็นแบบฟอร์มกลางสำหรับทุกประเภทเครื่องมือ",
    )
    workorder_type = models.CharField(
        max_length=50,
        choices=WorkOrder.WORKORDER_TYPE_CHOICES,
        verbose_name="ประเภทใบงาน",
    )
    description = models.TextField(blank=True, default="", verbose_name="คำอธิบาย")
    active = models.BooleanField(default=True, verbose_name="เปิดใช้งาน")
    version = models.PositiveIntegerField(default=1, verbose_name="เวอร์ชัน")
    # ตั้งเมื่อแบบฟอร์มนี้เคยถูกใช้กรอกผลแล้วและมีการแก้ไข -> สร้างเวอร์ชันใหม่แทนการเขียนทับ
    # เพื่อให้ผลที่กรอกไว้ในใบงานเก่ายังอ้างอิงรายการ/เกณฑ์ชุดเดิมได้ถูกต้อง
    replaced_by = models.OneToOneField(
        "self", null=True, blank=True, on_delete=models.SET_NULL,
        related_name="replaces", verbose_name="ถูกแทนที่ด้วยเวอร์ชัน",
    )
    created_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name="checklist_templates_created",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["workorder_type", "equipment_type", "name"]
        verbose_name = "แบบฟอร์ม Checklist"
        verbose_name_plural = "แบบฟอร์ม Checklist"
        indexes = [
            models.Index(fields=["workorder_type", "equipment_type", "active"]),
        ]

    def __str__(self):
        etype = self.equipment_type_display or "ทุกประเภท"
        return f"{self.name} v{self.version} [{self.get_workorder_type_display()} / {etype}]"

    @property
    def is_latest_version(self):
        return self.replaced_by_id is None

    @property
    def previous_version(self):
        """เวอร์ชันก่อนหน้าของแบบฟอร์มนี้ (ถ้ามี) — หาแบบ query แทนการพึ่ง reverse OneToOne accessor ตรงๆ เพื่อไม่ต้องดัก DoesNotExist."""
        return ChecklistTemplate.objects.filter(replaced_by=self).first()

    @property
    def equipment_type_list(self):
        return [v.strip() for v in self.equipment_type.split(",") if v.strip()]

    @property
    def equipment_type_display(self):
        return ", ".join(self.equipment_type_list)

    @classmethod
    def find_for(cls, workorder_type, equipment_type):
        """หาแบบฟอร์มที่ active ตรงกับประเภทเครื่องมือ (แบบฟอร์มหนึ่งผูกได้หลายชนิด) ก่อน แล้วค่อย fallback เป็นแบบฟอร์มกลาง."""
        qs = cls.objects.filter(workorder_type=workorder_type, active=True)
        if equipment_type:
            for tpl in qs.exclude(equipment_type=""):
                if equipment_type in tpl.equipment_type_list:
                    return tpl
        return qs.filter(equipment_type="").first()


class ChecklistTemplateItem(models.Model):
    RESULT_TYPE_CHOICES = [
        ("check", "ตรวจสอบ (ผ่าน/ไม่ผ่าน/N-A)"),
        ("measurement", "วัดค่า (เทียบค่ามาตรฐาน + ค่าเผื่อเบี่ยงเบน)"),
        ("text", "กรอกข้อความ/หมายเหตุ"),
    ]

    template = models.ForeignKey(
        ChecklistTemplate, on_delete=models.CASCADE, related_name="items"
    )
    order = models.PositiveIntegerField(default=0, verbose_name="ลำดับ")
    name = models.CharField(max_length=255, verbose_name="รายการตรวจสอบ")
    description = models.CharField(max_length=500, blank=True, default="", verbose_name="คำอธิบายเพิ่มเติม")
    result_type = models.CharField(
        max_length=20, choices=RESULT_TYPE_CHOICES, default="check",
        verbose_name="รูปแบบการกรอกผล",
    )
    unit = models.CharField(max_length=50, blank=True, default="", verbose_name="หน่วย")
    standard_value = models.DecimalField(
        max_digits=12, decimal_places=4, null=True, blank=True,
        verbose_name="ค่ามาตรฐาน/ค่าอ้างอิง",
    )
    tolerance = models.DecimalField(
        max_digits=12, decimal_places=4, null=True, blank=True,
        verbose_name="ค่าเผื่อเบี่ยงเบน (±)",
    )
    is_required = models.BooleanField(default=True, verbose_name="ต้องกรอก")

    class Meta:
        ordering = ["order", "id"]
        verbose_name = "รายการใน Checklist"
        verbose_name_plural = "รายการใน Checklist"

    def __str__(self):
        return f"{self.template_id}:{self.order} {self.name}"

    def evaluate(self, measured_value):
        """คืนค่า True/False/None ว่าผ่านเกณฑ์หรือไม่ สำหรับรายการแบบ measurement."""
        if measured_value is None or self.standard_value is None or self.tolerance is None:
            return None
        try:
            diff = abs(float(measured_value) - float(self.standard_value))
            return diff <= float(self.tolerance)
        except (TypeError, ValueError):
            return None


class WorkOrderChecklist(models.Model):
    """ผลการกรอก checklist ของใบงานหนึ่งใบ (แทนไฟล์ Excel/PDF ที่เคยแนบ)."""

    workorder = models.OneToOneField(
        WorkOrder, on_delete=models.CASCADE, related_name="checklist"
    )
    template = models.ForeignKey(
        ChecklistTemplate, on_delete=models.PROTECT, related_name="workorder_checklists"
    )
    completed_by = models.ForeignKey(
        settings.AUTH_USER_MODEL, null=True, blank=True, on_delete=models.SET_NULL,
        related_name="completed_checklists",
    )
    completed_at = models.DateTimeField(null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        verbose_name = "Checklist ของใบงาน"
        verbose_name_plural = "Checklist ของใบงาน"

    def __str__(self):
        return f"Checklist for WO#{self.workorder_id} ({self.template.name})"

    @property
    def overall_result(self):
        """'pass' | 'fail' | 'pending' — สรุปผลรวมของทุกรายการที่ต้องกรอก."""
        results_by_item = {r.template_item_id: r for r in self.results.all()}
        pending = False
        for item in self.template.items.all():
            r = results_by_item.get(item.id)
            if r is None or (item.is_required and r.is_pass is None and not r.note and not r.result_choice):
                if item.is_required:
                    pending = True
                continue
            if r.is_pass is False:
                return "fail"
        return "pending" if pending else "pass"


class WorkOrderChecklistResult(models.Model):
    RESULT_CHOICE_CHOICES = [
        ("pass", "ผ่าน"),
        ("fail", "ไม่ผ่าน"),
        ("na", "N/A"),
    ]

    checklist = models.ForeignKey(
        WorkOrderChecklist, on_delete=models.CASCADE, related_name="results"
    )
    template_item = models.ForeignKey(
        ChecklistTemplateItem, on_delete=models.PROTECT, related_name="results"
    )
    result_choice = models.CharField(
        max_length=10, choices=RESULT_CHOICE_CHOICES, blank=True, default="",
    )
    measured_value = models.DecimalField(
        max_digits=12, decimal_places=4, null=True, blank=True,
    )
    is_pass = models.BooleanField(null=True, blank=True)
    note = models.TextField(blank=True, default="")
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["template_item__order", "id"]
        unique_together = [("checklist", "template_item")]
        verbose_name = "ผลการกรอก Checklist รายข้อ"
        verbose_name_plural = "ผลการกรอก Checklist รายข้อ"

    def __str__(self):
        return f"{self.checklist_id}:{self.template_item_id} = {self.result_choice or self.measured_value}"

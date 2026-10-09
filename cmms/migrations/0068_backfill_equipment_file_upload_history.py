from django.db import migrations


def backfill_file_upload_history(apps, schema_editor):
    EquipmentProfileFile = apps.get_model("cmms", "EquipmentProfileFile")
    EquipmentHistory = apps.get_model("cmms", "EquipmentHistory")
    database = schema_editor.connection.alias

    history_entries = []
    for attachment in EquipmentProfileFile.objects.using(database).select_related(
        "equipment"
    ):
        section_label = (
            "รูปภาพ" if attachment.section == "image" else "เอกสาร/คู่มือ"
        )
        history_entries.append(
            EquipmentHistory(
                equipment_id=attachment.equipment.equipment_id,
                action_type="FILE_UPLOAD",
                action_by=attachment.uploaded_by or "ไม่ระบุ",
                action_at=attachment.uploaded_at,
                notes=f"อัปโหลด{section_label}: {attachment.display_name}",
            )
        )

    EquipmentHistory.objects.using(database).bulk_create(history_entries)


class Migration(migrations.Migration):
    dependencies = [
        ("cmms", "0067_alter_equipmenthistory_action_type"),
    ]

    operations = [
        migrations.RunPython(
            backfill_file_upload_history,
            reverse_code=migrations.RunPython.noop,
        ),
    ]

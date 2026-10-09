from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ("cmms", "0068_backfill_equipment_file_upload_history"),
    ]

    operations = [
        migrations.AlterField(
            model_name="equipmenthistory",
            name="action_type",
            field=models.CharField(
                choices=[
                    ("CREATE", "สร้างทะเบียน"),
                    ("UPDATE", "แก้ไขทะเบียน"),
                    ("DELETE", "ลบทะเบียน"),
                    ("FILE_UPLOAD", "อัปโหลดไฟล์"),
                    ("FILE_DELETE", "ลบไฟล์"),
                ],
                max_length=20,
                verbose_name="ประเภทการดำเนินการ",
            ),
        ),
    ]

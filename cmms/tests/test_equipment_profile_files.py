import os
import tempfile

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group, Permission
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.urls import reverse

from cmms.models import Equipment_list, EquipmentHistory, EquipmentProfileFile


class EquipmentProfileFileTests(TestCase):
    def setUp(self):
        self.equipment = Equipment_list.objects.create(
            equipment_id="MEDCMU-FILES-001",
            equipment_name_TH="เครื่องทดสอบ",
        )
        self.admin = get_user_model().objects.create_user(
            username="equipment-profile-admin",
            password="password",
            first_name="Test",
            last_name="Admin",
        )
        admin_group, _ = Group.objects.get_or_create(name="Admin")
        self.admin.groups.add(admin_group)
        self.admin.user_permissions.add(
            Permission.objects.get(codename="change_equipment_list")
        )
        self.client.force_login(self.admin)
        self.media_temp = tempfile.TemporaryDirectory()
        self.media_override = override_settings(MEDIA_ROOT=self.media_temp.name)
        self.media_override.enable()
        self.add_url = reverse(
            "equipment_profile_file_add",
            args=[self.equipment.equipment_id, EquipmentProfileFile.IMAGE],
        )

    def tearDown(self):
        self.media_override.disable()
        self.media_temp.cleanup()

    @staticmethod
    def image_upload(name="camera.png", data=b"\x89PNG\r\n\x1a\nsample"):
        return SimpleUploadedFile(name, data, content_type="image/png")

    @staticmethod
    def pdf_upload(name="manual.pdf"):
        return SimpleUploadedFile(name, b"%PDF-1.7 sample", content_type="application/pdf")

    def test_admin_can_upload_multiple_images_and_pdf_and_search_by_name(self):
        response = self.client.post(
            self.add_url,
            {
                "files": [
                    self.image_upload("front.png"),
                    self.image_upload("side.png"),
                    self.pdf_upload("image-notes.pdf"),
                ]
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            EquipmentProfileFile.objects.filter(
                equipment=self.equipment, section=EquipmentProfileFile.IMAGE
            ).count(),
            3,
        )
        self.assertEqual(
            set(
                EquipmentProfileFile.objects.filter(equipment=self.equipment)
                .values_list("uploaded_by", flat=True)
            ),
            {"Test Admin"},
        )
        self.assertEqual(
            EquipmentHistory.objects.filter(
                equipment_id=self.equipment.equipment_id,
                action_type="FILE_UPLOAD",
            ).count(),
            3,
        )
        profile_url = reverse(
            "equipment_profile", args=[self.equipment.equipment_id]
        )
        response = self.client.get(profile_url, {"tab": "images"})
        self.assertContains(response, "เพิ่มโดย Test Admin")
        self.assertEqual(
            list(response.context["profile_files"].values_list("display_name", flat=True)),
            ["image-notes.pdf", "side.png", "front.png"],
        )
        response = self.client.get(profile_url, {"tab": "images", "q": "side"})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "side.png")
        self.assertNotContains(response, "front.png")
        self.assertNotContains(response, "image-notes.pdf")
        self.assertContains(response, "แนบไฟล์")

    def test_document_menu_accepts_pdf_excel_and_image(self):
        upload_url = reverse(
            "equipment_profile_file_add",
            args=[self.equipment.equipment_id, EquipmentProfileFile.DOCUMENT],
        )
        response = self.client.post(
            upload_url,
            {
                "files": [
                    self.pdf_upload("service-manual.pdf"),
                    SimpleUploadedFile(
                        "checklist.xlsx", b"PK\x03\x04workbook", content_type="application/zip"
                    ),
                    self.image_upload("wiring.png"),
                ]
            },
        )

        self.assertEqual(response.status_code, 302)
        response = self.client.get(
            reverse("equipment_profile", args=[self.equipment.equipment_id]),
            {"tab": "documents"},
        )
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "service-manual.pdf")
        self.assertContains(response, "checklist.xlsx")
        self.assertContains(response, "wiring.png")
        history_response = self.client.get(
            reverse("equipment_history", args=[self.equipment.id])
        )
        self.assertEqual(history_response.status_code, 200)
        self.assertContains(history_response, "อัปโหลดไฟล์")
        self.assertContains(history_response, "อัปโหลดเอกสาร/คู่มือ")
        self.assertContains(history_response, "service-manual.pdf")
        self.assertContains(history_response, "checklist.xlsx")

    def test_admin_can_replace_rename_and_delete_file(self):
        response = self.client.post(
            self.add_url, {"files": [self.image_upload("original.png")]}
        )
        self.assertEqual(response.status_code, 302)
        attachment = EquipmentProfileFile.objects.get(equipment=self.equipment)
        old_file_name = attachment.file.name
        edit_url = reverse(
            "equipment_profile_file_edit",
            args=[
                self.equipment.equipment_id,
                EquipmentProfileFile.IMAGE,
                attachment.id,
            ],
        )
        response = self.client.post(
            edit_url,
            {
                "display_name": "replacement-name.png",
                "file": self.image_upload(
                    "replacement.png", b"\x89PNG\r\n\x1a\nreplacement"
                ),
            },
        )

        self.assertEqual(response.status_code, 302)
        attachment.refresh_from_db()
        self.assertEqual(attachment.display_name, "replacement-name.png")
        self.assertEqual(attachment.uploaded_by, "Test Admin")
        self.assertEqual(
            EquipmentHistory.objects.filter(
                equipment_id=self.equipment.equipment_id,
                action_type="FILE_UPLOAD",
            ).count(),
            2,
        )
        self.assertNotEqual(attachment.file.name, old_file_name)
        self.assertFalse(os.path.exists(os.path.join(self.media_temp.name, old_file_name)))
        self.assertTrue(os.path.exists(os.path.join(self.media_temp.name, attachment.file.name)))

        delete_url = reverse(
            "equipment_profile_file_delete",
            args=[
                self.equipment.equipment_id,
                EquipmentProfileFile.IMAGE,
                attachment.id,
            ],
        )
        response = self.client.post(delete_url)
        self.assertEqual(response.status_code, 302)
        self.assertFalse(EquipmentProfileFile.objects.filter(id=attachment.id).exists())
        self.assertFalse(os.path.exists(os.path.join(self.media_temp.name, attachment.file.name)))
        history_response = self.client.get(
            reverse("equipment_history", args=[self.equipment.id])
        )
        self.assertEqual(history_response.status_code, 200)
        self.assertContains(history_response, "ลบไฟล์")
        self.assertContains(history_response, "ลบรูปภาพ: replacement-name.png")
        deletion_record = EquipmentHistory.objects.get(
            equipment_id=self.equipment.equipment_id,
            action_type="FILE_DELETE",
        )
        self.assertEqual(deletion_record.action_by, "Test Admin")

    def test_non_admin_can_view_but_cannot_manage_files(self):
        attachment = EquipmentProfileFile.objects.create(
            equipment=self.equipment,
            section=EquipmentProfileFile.IMAGE,
            display_name="read-only.png",
            file=self.image_upload("read-only.png"),
        )
        viewer = get_user_model().objects.create_user(
            username="equipment-profile-viewer", password="password"
        )
        viewer_group, _ = Group.objects.get_or_create(name="Viewer")
        viewer.groups.add(viewer_group)
        self.client.force_login(viewer)
        self.client.raise_request_exception = False

        profile_url = reverse(
            "equipment_profile", args=[self.equipment.equipment_id]
        )
        response = self.client.get(profile_url, {"tab": "images"})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "read-only.png")
        self.assertContains(response, "เพิ่มโดย")
        self.assertContains(response, "ดูและค้นหาไฟล์ได้")
        self.assertNotContains(response, 'name="display_name"')
        self.assertNotContains(response, 'name="file"')
        self.assertContains(
            response,
            reverse(
                "equipment_profile_file_download",
                args=[self.equipment.equipment_id, attachment.id],
            ),
        )

        view_response = self.client.get(
            reverse(
                "equipment_profile_file_view",
                args=[self.equipment.equipment_id, attachment.id],
            )
        )
        self.assertEqual(view_response.status_code, 200)
        self.assertIn("inline", view_response["Content-Disposition"])
        self.assertEqual(
            b"".join(view_response.streaming_content),
            b"\x89PNG\r\n\x1a\nsample",
        )
        view_response.close()
        download_response = self.client.get(
            reverse(
                "equipment_profile_file_download",
                args=[self.equipment.equipment_id, attachment.id],
            )
        )
        self.assertEqual(download_response.status_code, 200)
        self.assertIn("attachment", download_response["Content-Disposition"])
        self.assertEqual(
            b"".join(download_response.streaming_content),
            b"\x89PNG\r\n\x1a\nsample",
        )
        download_response.close()

        for url, data in (
            (
                self.add_url,
                {"files": [self.image_upload("unauthorized.png")]},
            ),
            (
                reverse(
                    "equipment_profile_file_edit",
                    args=[
                        self.equipment.equipment_id,
                        EquipmentProfileFile.IMAGE,
                        attachment.id,
                    ],
                ),
                {"display_name": "edited.png"},
            ),
            (
                reverse(
                    "equipment_profile_file_delete",
                    args=[
                        self.equipment.equipment_id,
                        EquipmentProfileFile.IMAGE,
                        attachment.id,
                    ],
                ),
                {},
            ),
        ):
            with self.subTest(url=url):
                response = self.client.post(url, data)
                self.assertEqual(response.status_code, 403)

        self.assertTrue(EquipmentProfileFile.objects.filter(id=attachment.id).exists())
        self.assertFalse(
            EquipmentProfileFile.objects.filter(
                equipment=self.equipment, display_name="unauthorized.png"
            ).exists()
        )

    def test_equipment_edit_uploads_profile_images_and_documents_to_profile_storage(self):
        edit_url = reverse("edit_equipment", args=[self.equipment.id])
        response = self.client.get(edit_url)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "อัปโหลดรูปภาพโปรไฟล์")
        self.assertContains(response, 'name="profile_images"')
        self.assertContains(response, 'name="profile_documents"')

        response = self.client.post(
            edit_url,
            {
                "equipment_id": self.equipment.equipment_id,
                "profile_images": [
                    self.image_upload("profile-detail.png"),
                    self.image_upload("rear-panel.png"),
                ],
                "profile_documents": [
                    self.pdf_upload("service-guide.pdf"),
                    self.pdf_upload("maintenance-guide.pdf"),
                ],
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            set(
                EquipmentProfileFile.objects.filter(equipment=self.equipment)
                .values_list("section", "display_name")
            ),
            {
                (EquipmentProfileFile.IMAGE, "profile-detail.png"),
                (EquipmentProfileFile.IMAGE, "rear-panel.png"),
                (EquipmentProfileFile.DOCUMENT, "service-guide.pdf"),
                (EquipmentProfileFile.DOCUMENT, "maintenance-guide.pdf"),
            },
        )
        self.assertEqual(
            EquipmentHistory.objects.filter(
                equipment_id=self.equipment.equipment_id,
                action_type="FILE_UPLOAD",
            ).count(),
            4,
        )
        history_response = self.client.get(
            reverse("equipment_history", args=[self.equipment.id])
        )
        self.assertContains(history_response, "profile-detail.png")
        self.assertContains(history_response, "rear-panel.png")
        self.assertContains(history_response, "service-guide.pdf")
        self.assertContains(history_response, "maintenance-guide.pdf")

    def test_non_admin_cannot_upload_profile_files_from_equipment_edit(self):
        viewer = get_user_model().objects.create_user(
            username="equipment-profile-editor", password="password"
        )
        viewer.user_permissions.add(
            Permission.objects.get(codename="change_equipment_list")
        )
        self.client.force_login(viewer)
        edit_url = reverse("edit_equipment", args=[self.equipment.id])

        response = self.client.get(edit_url)
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "อัปโหลดรูปภาพโปรไฟล์")
        self.assertNotContains(response, 'name="profile_images"')
        self.assertNotContains(response, 'name="profile_documents"')

        self.client.raise_request_exception = False
        response = self.client.post(
            edit_url,
            {
                "equipment_id": self.equipment.equipment_id,
                "profile_images": [self.image_upload("unauthorized.png")],
            },
        )
        self.assertEqual(response.status_code, 403)
        self.assertFalse(
            EquipmentProfileFile.objects.filter(
                equipment=self.equipment, display_name="unauthorized.png"
            ).exists()
        )

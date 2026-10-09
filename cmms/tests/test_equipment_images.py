import os
import tempfile
from unittest.mock import Mock, patch

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Permission
from django.core.files.uploadedfile import SimpleUploadedFile
from django.db import IntegrityError
from django.test import TestCase, override_settings
from django.urls import reverse

from cmms.models import Equipment_list


class EquipmentImageTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(
            username="equipment-image-user", password="test-password"
        )
        permission = Permission.objects.get(codename="add_equipment_list")
        user.user_permissions.add(permission)
        self.client.force_login(user)

    def test_equipment_form_uses_upload_without_url_field(self):
        response = self.client.get(reverse("add_equipment"))

        self.assertEqual(response.status_code, 200)
        self.assertNotContains(response, 'name="equipment_photo"')
        self.assertContains(response, 'name="equipment_image"')

    def test_duplicate_equipment_id_is_rejected_with_visible_error(self):
        equipment_id = "MEDCMU-0000004"
        Equipment_list.objects.create(equipment_id=equipment_id)

        response = self.client.post(
            reverse("add_equipment"),
            {"equipment_id": equipment_id},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "รหัสนี้มีผู้ใช้แล้ว กรุณาใช้รหัสเครื่องอื่น")
        self.assertEqual(
            Equipment_list.objects.filter(equipment_id=equipment_id).count(),
            1,
        )

    def test_concurrent_duplicate_equipment_id_is_rejected_with_visible_error(self):
        equipment_id = "MEDCMU-0000005"
        with (
            patch("cmms.views.Equipment_list.objects.filter") as filter_mock,
            patch(
                "cmms.views.Equipment_list.objects.create",
                side_effect=IntegrityError("duplicate equipment_id"),
            ),
        ):
            filter_mock.side_effect = [
                Mock(exists=Mock(return_value=False)),
                Mock(exists=Mock(return_value=True)),
            ]
            response = self.client.post(
                reverse("add_equipment"),
                {"equipment_id": equipment_id},
            )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "รหัสนี้มีผู้ใช้แล้ว กรุณาใช้รหัสเครื่องอื่น")

    def test_upload_renames_image_and_exposes_view_and_download(self):
        with tempfile.TemporaryDirectory() as media_root:
            with override_settings(MEDIA_ROOT=media_root):
                upload = SimpleUploadedFile(
                    "original-name.png",
                    b"\x89PNG\r\n\x1a\n" + b"test-image",
                    content_type="image/png",
                )
                response = self.client.post(
                    reverse("add_equipment"),
                    {
                        "equipment_id": "MEDCMU-0000001",
                        "equipment_image": upload,
                    },
                )

                self.assertEqual(response.status_code, 302)
                equipment = Equipment_list.objects.get(
                    equipment_id="MEDCMU-0000001"
                )
                self.assertEqual(
                    equipment.equipment_image.name,
                    "equipment/MEDCMU-0000001.png",
                )
                self.assertTrue(
                    os.path.isfile(
                        os.path.join(media_root, equipment.equipment_image.name)
                    )
                )
                profile_response = self.client.get(
                    reverse("equipment_profile", args=[equipment.equipment_id])
                )
                self.assertEqual(profile_response.status_code, 200)
                self.assertNotContains(profile_response, "ยังไม่มีรูปภาพแนบ")
                self.assertContains(
                    profile_response,
                    reverse("equipment_image_view", args=[equipment.id]),
                )

                view_response = self.client.get(
                    reverse("equipment_image_view", args=[equipment.id])
                )
                self.assertEqual(view_response.status_code, 200)
                self.assertEqual(view_response["Content-Type"], "image/png")
                self.assertIn("inline", view_response["Content-Disposition"])
                self.assertEqual(
                    b"".join(view_response.streaming_content),
                    b"\x89PNG\r\n\x1a\ntest-image",
                )
                view_response.close()

                download_response = self.client.get(
                    reverse("equipment_image_download", args=[equipment.id])
                )
                self.assertEqual(download_response.status_code, 200)
                self.assertIn("attachment", download_response["Content-Disposition"])
                download_response.close()

    def test_master_data_has_english_equipment_images_entry(self):
        response = self.client.get(reverse("master_data"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Equipment Images")
        self.assertContains(response, reverse("equipment_images"))

    def test_service_request_staff_menu_requires_view_permission(self):
        response = self.client.get(reverse("master_data"))

        self.assertEqual(response.status_code, 200)
        self.assertNotContains(response, 'href="/service-requests"')

        permission = Permission.objects.get(codename="view_servicerequest")
        user = get_user_model().objects.get(username="equipment-image-user")
        user.user_permissions.add(permission)
        self.client.logout()
        self.client.force_login(user)

        response = self.client.get(reverse("master_data"))
        self.assertContains(response, 'href="/service-requests"')

    def test_missing_legacy_image_shows_placeholder_without_actions(self):
        equipment = Equipment_list.objects.create(
            equipment_id="MEDCMU-0000002",
            equipment_image="equipment/missing-file.png",
        )

        response = self.client.get(reverse("equipment_images"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "ยังไม่มีรูปภาพแนบ")
        self.assertNotContains(
            response, reverse("equipment_image_view", args=[equipment.id])
        )
        self.assertNotContains(
            response, reverse("equipment_image_download", args=[equipment.id])
        )
        self.assertEqual(
            self.client.get(
                reverse("equipment_image_view", args=[equipment.id])
            ).status_code,
            404,
        )
        self.assertEqual(
            self.client.get(
                reverse("equipment_image_download", args=[equipment.id])
            ).status_code,
            404,
        )

    def test_equipment_profile_without_image_shows_placeholder(self):
        Equipment_list.objects.create(equipment_id="MEDCMU-0000003")

        response = self.client.get(
            reverse("equipment_profile", args=["MEDCMU-0000003"])
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "ยังไม่มีรูปภาพแนบ")

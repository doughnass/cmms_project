from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.urls import reverse

from cmms.models import Equipment_list, EquipmentProfileFile


class EquipmentProfileFileCatalogTests(TestCase):
    def setUp(self):
        self.user = get_user_model().objects.create_user(
            username="profile-file-catalog-user", password="password"
        )
        self.client.force_login(self.user)
        self.equipment = Equipment_list.objects.create(
            equipment_id="MEDCMU-CATALOG-001",
            equipment_name_TH="เครื่องทดสอบรายการไฟล์",
            equipment_name_EN="File Catalog Equipment",
        )
        self.image_file = EquipmentProfileFile.objects.create(
            equipment=self.equipment,
            section=EquipmentProfileFile.IMAGE,
            file=SimpleUploadedFile(
                "catalog-image.png",
                b"\x89PNG\r\n\x1a\nsample",
                content_type="image/png",
            ),
            display_name="catalog-image.png",
            uploaded_by="Catalog User",
        )
        self.document_file = EquipmentProfileFile.objects.create(
            equipment=self.equipment,
            section=EquipmentProfileFile.DOCUMENT,
            file=SimpleUploadedFile(
                "catalog-manual.pdf",
                b"%PDF-1.7 sample",
                content_type="application/pdf",
            ),
            display_name="catalog-manual.pdf",
            uploaded_by="Catalog User",
        )

    def test_master_data_has_english_links_for_both_file_catalogs(self):
        response = self.client.get(reverse("master_data"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Equipment Profile Images")
        self.assertContains(
            response, reverse("equipment_profile_images")
        )
        self.assertContains(response, "Equipment Documents / Manuals")
        self.assertContains(
            response, reverse("equipment_profile_documents")
        )

    def test_image_and_document_catalogs_are_separate_and_searchable(self):
        image_response = self.client.get(reverse("equipment_profile_images"))
        self.assertEqual(image_response.status_code, 200)
        self.assertContains(image_response, "catalog-image.png")
        self.assertNotContains(image_response, "catalog-manual.pdf")
        self.assertContains(image_response, reverse(
            "equipment_profile_file_download",
            args=[self.equipment.equipment_id, self.image_file.id],
        ))

        document_response = self.client.get(
            reverse("equipment_profile_documents"),
            {"q": "manual"},
        )
        self.assertEqual(document_response.status_code, 200)
        self.assertContains(document_response, "catalog-manual.pdf")
        self.assertNotContains(document_response, "catalog-image.png")

        no_match_response = self.client.get(
            reverse("equipment_profile_documents"),
            {"q": "missing-file"},
        )
        self.assertContains(no_match_response, "No files match your search.")

    def test_catalog_exposes_manage_link_only_to_admin(self):
        response = self.client.get(reverse("equipment_profile_images"))
        self.assertNotContains(response, ">Manage</a>")

        admin_group, _ = Group.objects.get_or_create(name="Admin")
        self.user.groups.add(admin_group)
        response = self.client.get(reverse("equipment_profile_images"))
        self.assertContains(response, ">Manage</a>")

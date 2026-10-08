from django.test import TestCase

from cmms.models import MasterItem


class CreateAccountTests(TestCase):
    def test_department_suggestions_use_active_customer_master_items(self):
        MasterItem.objects.create(
            category="customers", code="active-unit", label="หน่วยงานที่ใช้งาน"
        )
        MasterItem.objects.create(
            category="customers",
            code="inactive-unit",
            label="หน่วยงานที่ปิดใช้งาน",
            active=False,
        )

        response = self.client.get("/create-account/")

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'list="department-options"')
        self.assertContains(response, '<option value="หน่วยงานที่ใช้งาน">')
        self.assertNotContains(response, '<option value="หน่วยงานที่ปิดใช้งาน">')

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Permission
from django.test import TestCase
from django.urls import reverse

from cmms.models import MasterItem, WorkOrder


class WorkOrderTypeFilterTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(
            username="workorder-filter-user", password="test-password"
        )
        user.user_permissions.add(
            Permission.objects.get(codename="view_workorder")
        )
        self.client.force_login(user)

        self.calibration_type = MasterItem.objects.create(
            category="workorder_type", code="calibration", label="สอบเทียบ"
        )
        self.maintenance_type = MasterItem.objects.create(
            category="workorder_type", code="maintenance", label="บำรุงรักษา"
        )

        self.calibration_order = WorkOrder.objects.create(
            title="ตรวจสอบเครื่องสอบเทียบ",
            workorder_type="other",
        )
        self.calibration_order.workorder_types.add(self.calibration_type)

        self.maintenance_order = WorkOrder.objects.create(
            title="บำรุงรักษาเครื่อง",
            workorder_type="other",
        )
        self.maintenance_order.workorder_types.add(self.maintenance_type)

        self.legacy_order = WorkOrder.objects.create(
            title="งานสอบเทียบแบบเก่า",
            workorder_type="calibration",
        )
        self.title_only_order = WorkOrder.objects.create(
            title="สอบเทียบ",
            workorder_type="other",
        )

    def test_filters_workorders_by_master_type_including_legacy_records(self):
        response = self.client.get(
            reverse("workorder_list"),
            {"work_type": f"master:{self.calibration_type.pk}"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertQuerySetEqual(
            response.context["workorders"].order_by("id"),
            [self.calibration_order, self.legacy_order, self.title_only_order],
        )
        self.assertContains(response, "ความต้องการ / ประเภทงาน")
        self.assertContains(response, "สอบเทียบ")
        self.assertContains(response, "บำรุงรักษา")

    def test_filters_legacy_type_values(self):
        response = self.client.get(
            reverse("workorder_list"),
            {"work_type": "legacy:calibration"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertQuerySetEqual(
            response.context["workorders"].order_by("id"),
            [self.calibration_order, self.legacy_order, self.title_only_order],
        )

    def test_matches_legacy_workorder_by_visible_title(self):
        response = self.client.get(
            reverse("workorder_list"),
            {"work_type": f"master:{self.calibration_type.pk}"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertIn(self.title_only_order, response.context["workorders"])

    def test_preserves_work_type_when_changing_status_filter(self):
        response = self.client.get(
            reverse("workorder_list"),
            {"work_type": f"master:{self.calibration_type.pk}", "status": "open"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(
            response,
            f"work_type=master%3A{self.calibration_type.pk}",
            html=False,
        )

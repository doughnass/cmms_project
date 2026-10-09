from django.contrib.auth import get_user_model
from django.test import TestCase
from django.urls import reverse

from cmms.models import Equipment_list, MasterItem, WorkOrder


class EquipmentHistoryTypeFilterTests(TestCase):
    def setUp(self):
        user = get_user_model().objects.create_user(
            username="equipment-history-filter-user", password="password"
        )
        self.client.force_login(user)
        self.equipment = Equipment_list.objects.create(
            equipment_id="MEDCMU-HISTORY-001"
        )
        self.repair_type = MasterItem.objects.create(
            category="workorder_type", code="repair", label="ซ่อมแซม"
        )
        self.maintenance_type = MasterItem.objects.create(
            category="workorder_type", code="maintenance", label="บำรุงรักษา"
        )
        self.emergency_type = MasterItem.objects.create(
            category="workorder_type", code="breakdown", label="เสียฉุกเฉิน"
        )
        self.repair_order = WorkOrder.objects.create(
            equipment=self.equipment, title="เปลี่ยนอะไหล่", workorder_type="other"
        )
        self.repair_order.workorder_types.add(self.repair_type)
        self.emergency_order = WorkOrder.objects.create(
            equipment=self.equipment,
            title="เครื่องหยุดทำงาน",
            workorder_type="other",
        )
        self.emergency_order.workorder_types.add(
            self.repair_type, self.emergency_type
        )
        self.maintenance_order = WorkOrder.objects.create(
            equipment=self.equipment,
            title="ตรวจเช็กประจำปี",
            workorder_type="maintenance",
        )

    def test_history_shows_type_filters_and_resolves_multiple_work_order_types(self):
        response = self.client.get(
            reverse("equipment_history", args=[self.equipment.id])
        )

        self.assertEqual(response.status_code, 200)
        filters = {
            item["code"]: item for item in response.context["work_order_type_filters"]
        }
        self.assertEqual(filters["repair"]["label"], "ซ่อมแซม")
        self.assertEqual(filters["repair"]["count"], 2)
        self.assertEqual(filters["maintenance"]["count"], 1)
        self.assertEqual(filters["breakdown"]["label"], "เสียฉุกเฉิน")
        self.assertEqual(filters["breakdown"]["count"], 1)
        orders_by_id = {
            work_order.id: work_order
            for work_order in response.context["work_orders"]
        }
        self.assertEqual(
            orders_by_id[self.emergency_order.id].history_type_codes,
            {"repair", "breakdown"},
        )
        self.assertContains(response, "ให้คำแนะนำ")
        self.assertContains(response, "ติดตั้ง")
        self.assertContains(response, "อื่น ๆ")
        self.assertContains(response, 'data-filter-type="breakdown"')

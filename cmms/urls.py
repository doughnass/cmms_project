from django.contrib.auth import views as auth_views
from django.urls import path

from cmms import views
from cmms import api_availability
from cmms.forms import ThaiAuthenticationForm

from . import views_admin, views_holiday, views_reports

urlpatterns = [
    path("", views.index, name="index"),
    path("create_request", views.create_request, name="create_request"),
    path("equipment_list/", views.equipment_list, name="equipment_list"),
    path("equipment-images/", views.equipment_images, name="equipment_images"),
    path(
        "equipment-image/<int:equipment_list_id>/",
        views.equipment_image_view,
        name="equipment_image_view",
    ),
    path(
        "equipment-image/<int:equipment_list_id>/download/",
        views.equipment_image_download,
        name="equipment_image_download",
    ),
    path("api/equipment_list/", views.api_equipment_list, name="api_equipment_list"),
    path("add_equipment/", views.add_equipment, name="add_equipment"),
    path(
        "edit_equipment/<equipment_list_id>/",
        views.edit_equipment,
        name="edit_equipment",
    ),
    path(
        "edit_equipment_by_code/<str:code>/",
        views.edit_equipment_by_code,
        name="edit_equipment_by_code",
    ),
    path(
        "equipment_history/<equipment_list_id>/",
        views.equipment_history,
        name="equipment_history",
    ),
    path(
        "equipment_profile/<str:code>/",
        views.equipment_profile,
        name="equipment_profile",
    ),
    path(
        "equipment_profile/<str:code>/files/<str:section>/add/",
        views.equipment_profile_file_add,
        name="equipment_profile_file_add",
    ),
    path(
        "equipment_profile/<str:code>/files/<int:file_id>/view/",
        views.equipment_profile_file_view,
        name="equipment_profile_file_view",
    ),
    path(
        "equipment_profile/<str:code>/files/<int:file_id>/download/",
        views.equipment_profile_file_view,
        {"download": True},
        name="equipment_profile_file_download",
    ),
    path(
        "equipment-profile-files/images/",
        views.equipment_profile_files_list,
        {"section": "image"},
        name="equipment_profile_images",
    ),
    path(
        "equipment-profile-files/documents/",
        views.equipment_profile_files_list,
        {"section": "document"},
        name="equipment_profile_documents",
    ),
    path(
        "equipment_profile/<str:code>/files/<str:section>/<int:file_id>/edit/",
        views.equipment_profile_file_edit,
        name="equipment_profile_file_edit",
    ),
    path(
        "equipment_profile/<str:code>/files/<str:section>/<int:file_id>/delete/",
        views.equipment_profile_file_delete,
        name="equipment_profile_file_delete",
    ),
    path(
        "equipment_profile_by_code/<str:code>/",
        views.equipment_profile_by_code,
        name="equipment_profile_by_code",
    ),
    path(
        "delete_equipment/<equipment_list_id>/",
        views.delete_equipment,
        name="delete_equipment",
    ),
    path(
        "delete_equipment_by_code/<str:code>/",
        views.delete_equipment_by_code,
        name="delete_equipment_by_code",
    ),
    path(
        "import_equipment_list",
        views.import_equipment_list,
        name="import_equipment_list",
    ),
    path(
        "equipment_download_template",
        views.equipment_download_template,
        name="equipment_download_template",
    ),
    path(
        "bulk_edit_equipment",
        views.bulk_edit_equipment,
        name="bulk_edit_equipment",
    ),
    path(
        "export_equipment_list",
        views.export_equipment_list,
        name="export_equipment_list",
    ),
    path(
        "maintenance/schedule/", views.maintenance_schedule, name="maintenance_schedule"
    ),
    path(
        "maintenance/capacities/",
        views.maintenance_capacity_list,
        name="maintenance_capacity_list",
    ),
    path(
        "maintenance/capacities/add/",
        views.maintenance_capacity_create,
        name="maintenance_capacity_create",
    ),
    path(
        "maintenance/capacities/delete/<int:cap_id>/",
        views.maintenance_capacity_delete,
        name="maintenance_capacity_delete",
    ),
    path(
        "maintenance/capacities/edit/<int:cap_id>/",
        views.maintenance_capacity_edit,
        name="maintenance_capacity_edit",
    ),
    path("maintenance/technicians/", views.technicians_list, name="technicians_list"),
    path(
        "maintenance/technicians/add/",
        views.technician_create,
        name="technician_create",
    ),
    path(
        "maintenance/technicians/edit/<int:tech_id>/",
        views.technician_edit,
        name="technician_edit",
    ),
    path(
        "maintenance/technicians/delete/<int:tech_id>/",
        views.technician_delete,
        name="technician_delete",
    ),
    path(
        "maintenance/technicians/availability/",
        views.technician_availability_manage,
        name="technician_availability_manage",
    ),
    path(
        "maintenance/technicians/availability/v2/",
        views.technician_availability_manage_v2,
        name="technician_availability_manage_v2",
    ),
    path(
        "maintenance/technicians/availability/report/",
        views.technician_availability_report,
        name="technician_availability_report",
    ),
    path(
        "maintenance/daily-capacity/",
        views.daily_capacity_management,
        name="daily_capacity_management",
    ),
    path(
        "maintenance/capacity/summary/",
        views.capacity_summary,
        name="capacity_summary",
    ),
    path("maintenance/shift-schedule/", views.shift_schedule, name="shift_schedule"),
    path(
        "maintenance/shift-schedule/report/",
        views.shift_schedule_report,
        name="shift_schedule_report",
    ),
    path(
        "api/shift-schedule/save/",
        views.api_save_shift_schedule,
        name="api_save_shift_schedule",
    ),
    path(
        "api/shift-schedule/<int:schedule_id>/delete/",
        views.api_delete_shift_schedule,
        name="api_delete_shift_schedule",
    ),
    path(
        "api/maintenance/technicians/",
        views.api_technicians_by_date,
        name="api_technicians_by_date",
    ),
    path(
        "api/maintenance/availability/",
        api_availability.api_maintenance_availability,
        name="api_maintenance_availability",
    ),
    path(
        "api/maintenance/daily-summary/",
        views.api_daily_summary,
        name="api_daily_summary",
    ),
    path(
        "api/select2/equipment/",
        views.api_equipment_for_select2,
        name="api_select2_equipment",
    ),
    path(
        "api/select2/skills/",
        views.api_master_skills_for_select2,
        name="api_select2_skills",
    ),
    path(
        "maintenance/appointments/",
        views.maintenance_appointments_list,
        name="maintenance_appointments_list",
    ),
    path(
        "maintenance/technician-schedule/",
        views.maintenance_technician_schedule,
        name="maintenance_technician_schedule",
    ),
    path(
        "maintenance/appointments/delete/<int:appt_id>/",
        views.maintenance_appointment_delete,
        name="maintenance_appointment_delete",
    ),
    path(
        "maintenance/appointments/edit/<int:appt_id>/",
        views.maintenance_appointment_edit,
        name="maintenance_appointment_edit",
    ),
    path(
        "api/maintenance/check_capacity/",
        views.api_check_capacity,
        name="api_check_capacity",
    ),
    # Equipment maintenance schedule
    path(
        "maintenance/equipment-schedule/",
        views.equipment_maintenance_schedule,
        name="equipment_maintenance_schedule",
    ),
    path(
        "api/maintenance/equipment/due/",
        views.api_equipment_due_list,
        name="api_equipment_due_list",
    ),
    path(
        "api/maintenance/equipment/schedule/",
        views.api_schedule_equipment,
        name="api_schedule_equipment",
    ),
    path(
        "api/maintenance/equipment/bulk-schedule/",
        views.api_bulk_schedule_equipment,
        name="api_bulk_schedule_equipment",
    ),
    path(
        "api/maintenance/appointments/auto-assign/",
        views.api_auto_assign_appointment,
        name="api_auto_assign_appointment",
    ),
    path(
        "api/maintenance/equipment/create-service-request/",
        views.api_create_service_request,
        name="api_create_service_request",
    ),
    path(
        "api/maintenance/equipment/bulk-create-service-request/",
        views.api_bulk_create_service_request,
        name="api_bulk_create_service_request",
    ),
    path(
        "api/maintenance/equipment/has-service-requests/",
        views.api_service_requests_by_equipment,
        name="api_equipment_has_service_requests",
    ),
    # Day schedule APIs
    path(
        "api/equipment/maintenance-schedule/",
        views.api_equipment_by_date,
        name="api_equipment_by_date",
    ),
    path(
        "api/stations/availability/",
        views.api_stations_availability,
        name="api_stations_availability",
    ),
    path(
        "api/technicians/availability/",
        views.api_technicians_availability,
        name="api_technicians_availability",
    ),
    path(
        "api/technicians/list/",
        views.api_technicians_list,
        name="api_technicians_list",
    ),
    path(
        "api/maintenance/equipment/update-estimated-hours/",
        views.api_update_equipment_estimated_hours,
        name="api_update_equipment_estimated_hours",
    ),
    # customer import/export routes
    path(
        "import_customer_list", views.import_customer_list, name="import_customer_list"
    ),
    path(
        "customer_list_export", views.export_customer_list, name="customer_list_export"
    ),
    path("customer_list", views.customer_list, name="customer_list"),
    path("add_customer", views.add_customer, name="add_customer"),
    path("edit_customer/<customer_list_id>", views.edit_customer, name="edit_customer"),
    path(
        "delete_customer/<customer_list_id>",
        views.delete_customer,
        name="delete_customer",
    ),
    path("master_data", views.master_data, name="master_data"),
    # master data management
    path("master/", views.master_list, name="master_list"),
    path(
        "master/category/<path:category>",
        views.master_list,
        name="master_list_by_category",
    ),
    path("master/manage/", views.master_tree_manage, name="master_manage"),
    path(
        "master/manage/<path:category>",
        views.master_tree_manage,
        name="master_manage_by_category",
    ),
    # AJAX API for master CRUD
    path("master/api/create/", views.master_api_create, name="master_api_create"),
    path(
        "master/api/update/<int:pk>/", views.master_api_update, name="master_api_update"
    ),
    path(
        "master/api/delete/<int:pk>/", views.master_api_delete, name="master_api_delete"
    ),
    path("master/api/get/<int:pk>/", views.master_api_get, name="master_api_get"),
    path("master/api/parents/", views.master_api_parents, name="master_api_parents"),
    path("master/api/reorder/", views.master_api_reorder, name="master_api_reorder"),
    path("master/api/undo/", views.master_api_undo, name="master_api_undo"),
    path("master/api/redo/", views.master_api_redo, name="master_api_redo"),
    path("master/add/", views.master_create, name="master_create"),
    path("master/edit/<int:pk>/", views.master_edit, name="master_edit"),
    path("master/delete/<int:pk>/", views.master_delete, name="master_delete"),
    path("master/import/", views.master_import_excel, name="master_import_excel"),
    path(
        "master/template/download/",
        views.master_download_template,
        name="master_download_template",
    ),
    # Master Category management (in-app)
    path("master/categories/", views.category_list, name="category_list"),
    path("master/categories/add/", views.category_create, name="category_create"),
    path(
        "master/categories/edit/<int:cat_id>/",
        views.category_edit,
        name="category_edit",
    ),
    path(
        "master/categories/delete/<int:cat_id>/",
        views.category_delete,
        name="category_delete",
    ),
    # developer tools: template inspection (staff only)
    path("tools/templates/", views.templates_index, name="templates_index"),
    path(
        "tools/templates/download/csv/",
        views.templates_index_csv,
        name="templates_index_csv",
    ),
    path(
        "tools/templates/download/csv/<str:group_key>/",
        views.templates_index_group_csv,
        name="templates_index_group_csv",
    ),
    path("tools/templates/source/", views.template_source, name="template_source"),
    path("tools/templates/render/", views.template_render, name="template_render"),
    path("reports", views.reports, name="reports"),
    path("about", views.about, name="about"),
    path("favicon.ico", views.favicon),
    path(
        "api/check_equipment_id",
        views.api_check_equipment_id,
        name="api_check_equipment_id",
    ),
    path(
        "api/search/equipment", views.api_search_equipment, name="api_search_equipment"
    ),
    path("api/search/users", views.api_search_users, name="api_search_users"),
    path("api/master/items", views.api_master_items, name="api_master_items"),
    path("service-terms", views.service_terms_page, name="service_terms_page"),
    path("service-terms/edit", views.service_terms_edit, name="service_terms_edit"),
    path("api/dashboard_stats/", views.api_dashboard_stats, name="api_dashboard_stats"),
    # work orders
    path("workorders", views.workorder_list, name="workorder_list"),
    path("workorders/create", views.workorder_create, name="workorder_create"),
    path("workorders/<int:wo_id>", views.workorder_detail, name="workorder_detail"),
    path("workorders/<int:wo_id>/attachments/<int:att_id>/delete", views.delete_workorder_attachment, name="delete_workorder_attachment"),
    path("workorders/<int:wo_id>/checklist/", views.workorder_checklist_fill, name="workorder_checklist_fill"),
    # checklist templates (แบบฟอร์ม PM/สอบเทียบ แก้ไขได้)
    path("checklist-templates/", views.checklist_templates, name="checklist_templates"),
    path("checklist-templates/add/", views.checklist_template_form, name="checklist_template_create"),
    path("checklist-templates/<int:pk>/edit/", views.checklist_template_form, name="checklist_template_edit"),
    path("checklist-templates/<int:pk>/delete/", views.checklist_template_delete, name="checklist_template_delete"),
    # service requests (public + officer)
    path(
        "service-request/create",
        views.service_request_create,
        name="service_request_create",
    ),
    path("service-requests", views.service_request_list, name="service_request_list"),
    path("my-service-requests", views.my_service_requests, name="my_service_requests"),
    path(
        "service-requests/<int:sr_id>",
        views.service_request_detail,
        name="service_request_detail",
    ),
    path(
        "service-requests/<int:sr_id>/cancel",
        views.service_request_cancel,
        name="service_request_cancel",
    ),
    path(
        "service-requests/<int:sr_id>/delete",
        views.service_request_delete,
        name="service_request_delete",
    ),
    path(
        "service-requests/deleted",
        views.deleted_requests_report,
        name="service_requests_deleted",
    ),
    path(
        "service-requests/trash",
        views.service_request_trash,
        name="service_request_trash",
    ),
    path(
        "service-requests/<int:sr_id>/restore",
        views.service_request_restore,
        name="service_request_restore",
    ),
    path(
        "service-requests/<int:sr_id>/purge",
        views.service_request_purge,
        name="service_request_purge",
    ),
    # authentication
    path(
        "login/",
        auth_views.LoginView.as_view(
            template_name="login.html", authentication_form=ThaiAuthenticationForm
        ),
        name="login",
    ),
    path("logout/", auth_views.LogoutView.as_view(next_page="index"), name="logout"),
    path("create-account/", views.create_account, name="create_account"),
    path(
        "accounts/approvals/",
        views.account_approvals,
        name="account_approvals",
    ),
    path(
        "accounts/approvals/<int:user_id>/approve/",
        views.account_approve,
        name="account_approve",
    ),
    path(
        "accounts/approvals/<int:user_id>/reject/",
        views.account_reject,
        name="account_reject",
    ),
    path(
        "accounts/manage/",
        views.user_management,
        name="user_management",
    ),
    path(
        "accounts/manage/<int:user_id>/update/",
        views.user_update_access,
        name="user_update_access",
    ),
    path("profile/", views.profile, name="profile"),
    path("profile/edit/", views.edit_profile, name="edit_profile"),
    path("profile/change-password/", views.change_password, name="change_password"),
    # multi-account management
    path("accounts/", views.multi_accounts, name="multi_accounts"),
    path("accounts/add/", views.add_account, name="add_account"),
    path("accounts/remove/", views.remove_account, name="remove_account"),
    path("accounts/switch/", views.switch_account, name="switch_account"),
    path(
        "admin/copied-masteritems/",
        views_admin.copied_masteritems_report,
        name="copied_masteritems_report",
    ),
    # Holiday API
    path("api/holidays/", views_holiday.holiday_list, name="api_holiday_list"),
    path(
        "api/holidays/create/", views_holiday.holiday_create, name="api_holiday_create"
    ),
    # Smart appointment availability API
    path(
        "api/appointment/available-dates/",
        views.api_available_dates_for_equipment,
        name="api_available_dates_for_equipment",
    ),
    path(
        "api/holidays/<int:holiday_id>/update/",
        views_holiday.holiday_update,
        name="api_holiday_update",
    ),
    path(
        "api/holidays/<int:holiday_id>/delete/",
        views_holiday.holiday_delete,
        name="api_holiday_delete",
    ),
    path("api/holidays/check/", views_holiday.holiday_check, name="api_holiday_check"),
    path(
        "api/holidays/import-government/",
        views_holiday.holiday_import_government,
        name="api_holiday_import_government",
    ),
    path(
        "api/holidays/sync-from-api/",
        views_holiday.holiday_sync_from_api,
        name="api_holiday_sync_from_api",
    ),
    # ── ทะเบียนอะไหล่คงคลัง (Spare Parts Inventory) ──
    path("spare-parts/", views.spare_part_list, name="spare_part_list"),
    path("spare-parts/create/", views.spare_part_create, name="spare_part_create"),
    path("spare-parts/<int:pk>/update/", views.spare_part_update, name="spare_part_update"),
    path("spare-parts/<int:pk>/delete/", views.spare_part_delete, name="spare_part_delete"),
    path("spare-parts/<int:pk>/adjust/", views.spare_part_adjust_stock, name="spare_part_adjust_stock"),
    path("api/spare-parts/search/", views.api_spare_parts_search, name="api_spare_parts_search"),
    
    # ── รายงาน KPI (KPI Reports) ──
    path("reports/kpi-dashboard/", views_reports.kpi_dashboard, name="kpi_dashboard"),
    path("reports/kpi-print/", views_reports.kpi_report_print, name="kpi_report_print"),
    path("reports/kpi-comparison/", views_reports.kpi_comparison, name="kpi_comparison"),
    path("api/kpi-data/", views_reports.api_kpi_data, name="api_kpi_data"),
    path("api/workorders-by-type/", views_reports.api_workorders_by_type, name="api_workorders_by_type"),
]

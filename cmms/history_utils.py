"""
Utility functions for logging equipment registration and edit history.
ฟังก์ชันสำหรับบันทึกประวัติการขึ้นทะเบียนและแก้ไขอุปกรณ์
"""

from django.utils import timezone
from .models import Equipment_list, EquipmentHistory


def get_client_ip(request):
    """Extract client IP address from request"""
    x_forwarded_for = request.META.get('HTTP_X_FORWARDED_FOR')
    if x_forwarded_for:
        ip = x_forwarded_for.split(',')[0]
    else:
        ip = request.META.get('REMOTE_ADDR')
    return ip


def get_user_display_name(user):
    """Get display name for user (full name or username)"""
    if not user or not user.is_authenticated:
        return 'Anonymous'
    return user.get_full_name() or user.username


def log_equipment_history(equipment_id, action_type, user, request=None, old_data=None, new_data=None, notes=''):
    """
    Log equipment registration or edit action to history.
    
    Args:
        equipment_id: Equipment ID (MEDCMU-XXXXX)
        action_type: 'CREATE', 'UPDATE', or 'DELETE'
        user: Django User object
        request: Django request object (for IP address)
        old_data: Dict of old field values (for UPDATE)
        new_data: Dict of new field values (for CREATE/UPDATE)
        notes: Optional notes
    
    Returns:
        EquipmentHistory instance
    """
    changed_fields = []
    old_values = {}
    new_values = {}
    
    # For UPDATE actions, calculate what changed
    if action_type == 'UPDATE' and old_data and new_data:
        for field, new_value in new_data.items():
            old_value = old_data.get(field)
            # Compare values (handle None, empty strings, dates, etc.)
            if str(old_value) != str(new_value):
                changed_fields.append(field)
                old_values[field] = old_value
                new_values[field] = new_value
    elif action_type == 'CREATE' and new_data:
        # For CREATE, all non-empty fields are "new"
        for field, value in new_data.items():
            if value not in [None, '', []]:
                changed_fields.append(field)
                new_values[field] = value
    
    history = EquipmentHistory.objects.create(
        equipment_id=equipment_id,
        action_type=action_type,
        action_by=get_user_display_name(user),
        ip_address=get_client_ip(request) if request else None,
        changed_fields=changed_fields,
        old_values=old_values,
        new_values=new_values,
        notes=notes
    )
    
    return history


def get_equipment_history(equipment_id, limit=None):
    """
    Get history for a specific equipment.
    
    Args:
        equipment_id: Equipment ID
        limit: Optional limit on number of records
    
    Returns:
        QuerySet of EquipmentHistory objects
    """
    history = EquipmentHistory.objects.filter(equipment_id=equipment_id).order_by('-action_at')
    if limit:
        history = history[:limit]
    return history


def format_field_name(field_name):
    """Convert field name to Thai display name"""
    field_labels = {
        'equipment_id': 'รหัสเครื่อง',
        'equipment_code': 'Equipment Master List (EML)',
        'equipment_name_EN': 'ชื่อเครื่อง (English)',
        'equipment_name_TH': 'ชื่อเครื่อง (ไทย)',
        'equipment_brand': 'ยี่ห้อ/ผู้ผลิต',
        'equipment_model': 'รุ่น (Model)',
        'equipment_sn': 'Serial Number',
        'equipment_gov': 'เลขครุภัณฑ์',
        'equipment_price': 'ราคา',
        'equipment_photo': 'รูปภาพ',
        'equipment_image': 'ไฟล์รูปภาพ',
        'equipment_type': 'ประเภท',
        'equipment_life': 'อายุการใช้งาน',
        'equipment_waranty_date': 'วันที่เริ่มรับประกัน',
        'equipment_waranty_due': 'วันที่หมดประกัน',
        'equipment_distributor_name': 'บริษัทผู้จำหน่าย',
        'equipment_distributor_tel': 'เบอร์ติดต่อผู้จำหน่าย',
        'equipment_pm_fq': 'ความถี่การบำรุงรักษา',
        'equipment_pm_due': 'ครบกำหนดการบำรุงรักษา',
        'equipment_cal_fq': 'ความถี่การสอบเทียบ',
        'equipment_cal_due': 'ครบกำหนดการสอบเทียบ',
        'requires_pm': 'ต้องการ PM',
        'requires_cal': 'ต้องการ CAL',
        'equipment_owner_customer': 'หน่วยงานเจ้าของ',
        'equipment_user_customer': 'หน่วยงานผู้ใช้',
        'equipment_service_provider': 'ผู้ดูแล/ผู้ให้บริการ',
        'equipment_register_username': 'ผู้ขอขึ้นทะเบียน',
        'equipment_register_adminname': 'ผู้ขึ้นทะเบียน',
        'equipment_note': 'หมายเหตุ',
    }
    return field_labels.get(field_name, field_name)


def get_field_value_from_equipment(equipment, field_name):
    """Extract field value from Equipment_list object"""
    value = getattr(equipment, field_name, None)
    
    # Format dates
    if hasattr(value, 'isoformat'):
        return value.isoformat()
    
    # Handle None/empty
    if value is None or value == '':
        return ''
    
    return str(value)


def capture_equipment_snapshot(equipment):
    """
    Capture current state of all equipment fields.
    
    Args:
        equipment: Equipment_list instance
    
    Returns:
        Dict of field_name: value
    """
    fields_to_track = [
        'equipment_id', 'equipment_code', 'equipment_name_EN', 'equipment_name_TH',
        'equipment_brand', 'equipment_model', 'equipment_sn', 'equipment_gov',
        'equipment_price', 'equipment_photo', 'equipment_image', 'equipment_type', 'equipment_life',
        'equipment_waranty_date', 'equipment_waranty_due',
        'equipment_distributor_name', 'equipment_distributor_tel',
        'equipment_pm_fq', 'equipment_pm_due',
        'equipment_cal_fq', 'equipment_cal_due',
        'requires_pm', 'requires_cal',
        'equipment_owner_customer', 'equipment_user_customer',
        'equipment_service_provider',
        'equipment_register_username', 'equipment_register_adminname',
        'equipment_note',
    ]
    
    snapshot = {}
    for field in fields_to_track:
        snapshot[field] = get_field_value_from_equipment(equipment, field)
    
    return snapshot

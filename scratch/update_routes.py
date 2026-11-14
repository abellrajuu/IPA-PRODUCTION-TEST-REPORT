import sys, os

path = r'C:\Users\Admin\OneDrive\Desktop\WITH STERP\PDN TEST REPORT\backend_python.py'
with open(path, 'r', encoding='utf-8') as f:
    content = f.read()

target = """SYSTEM_ROUTES = {
    "/": "portal.html",
    "/portal": "portal.html",

    # Documentation routes hidden for now per request:
    # "/job-traveller-card": "job_traveller_card.html",
    # "/inprocess-register": "inprocess.html",
    # "/inprocess": "inprocess.html",

    "/belt-scale": "belt_scale.html",
    "/batching-system": "batching_system.html",
    "/remote-indicator": "remote_indicator.html",
    "/digital-indicator": "digital_indicator.html",
    "/weigh-feeder": "weigh_feeder.html",
    "/crane-scale": "crane_scale.html",
    "/signal-conditioner": "signal_conditioner.html",
    "/trip-safe": "trip_safe.html",
    "/vibration-switch": "vibration_switch.html",
    "/acc-mv": "acc_mv.html",
    "/acc-charge": "acc_charge.html",
    "/vibration-meter": "vibration_meter.html",
    "/charge-amplifier": "charge_amplifier.html",
    "/odd-system": "odd_system.html",
    "/dd-system": "dd_system.html",
    "/loss-in-weigh-feeder": "loss_in_weigh_feeder.html",
    "/misc-report": "misc_report.html",
    "/misc_report": "misc_report.html",
}"""

replacement = """SYSTEM_ROUTES = {
    "/": "portal.html",
    "/portal": "portal.html",

    "/belt-scale": "belt_scale.html",
    "/belt_scale": "belt_scale.html",
    "/batching-system": "batching_system.html",
    "/batching_system": "batching_system.html",
    "/remote-indicator": "remote_indicator.html",
    "/remote_indicator": "remote_indicator.html",
    "/digital-indicator": "digital_indicator.html",
    "/digital_indicator": "digital_indicator.html",
    "/weigh-feeder": "weigh_feeder.html",
    "/weigh_feeder": "weigh_feeder.html",
    "/crane-scale": "crane_scale.html",
    "/crane_scale": "crane_scale.html",
    "/signal-conditioner": "signal_conditioner.html",
    "/signal_conditioner": "signal_conditioner.html",
    "/trip-safe": "trip_safe.html",
    "/trip_safe": "trip_safe.html",
    "/vibration-switch": "vibration_switch.html",
    "/vibration_switch": "vibration_switch.html",
    "/acc-mv": "acc_mv.html",
    "/acc_mv": "acc_mv.html",
    "/acc-charge": "acc_charge.html",
    "/acc_charge": "acc_charge.html",
    "/vibration-meter": "vibration_meter.html",
    "/vibration_meter": "vibration_meter.html",
    "/charge-amplifier": "charge_amplifier.html",
    "/charge_amplifier": "charge_amplifier.html",
    "/odd-system": "odd_system.html",
    "/odd_system": "odd_system.html",
    "/dd-system": "dd_system.html",
    "/dd_system": "dd_system.html",
    "/loss-in-weigh-feeder": "loss_in_weigh_feeder.html",
    "/loss_in_weigh_feeder": "loss_in_weigh_feeder.html",
    "/job-traveller-card": "job_traveller_card.html",
    "/job_traveller_card": "job_traveller_card.html",
    "/inprocess": "inprocess.html",
    "/inprocess_register": "inprocess.html",
    "/misc-report": "misc_report.html",
    "/misc_report": "misc_report.html",
}"""

target_c = target.replace('\r\n', '\n')
content_c = content.replace('\r\n', '\n')
if target_c in content_c:
    new_content = content_c.replace(target_c, replacement)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('SUCCESS: Updated SYSTEM_ROUTES in backend_python.py')
else:
    print('Target not found in content!')

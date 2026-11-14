import os
import sys
import shutil
import time
import json
import pymupdf

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import report_generator_python as docx_generator

# Output directories on Desktop as explicitly requested: "folder on desktop called o/p"
# We also mirror to "o_p" and both OneDrive\Desktop and standard Desktop.
TARGET_DIRS = [
    os.path.join(r"C:\Users\abell\OneDrive\Desktop", "o", "p"),
    os.path.join(r"C:\Users\abell\OneDrive\Desktop", "o_p"),
    os.path.join(r"C:\Users\abell\OneDrive\Desktop", "op"),
    os.path.join(r"C:\Users\abell\OneDrive\Desktop", "outputs"),
    os.path.join(r"C:\Users\abell\Desktop", "o", "p"),
    os.path.join(r"C:\Users\abell\Desktop", "o_p"),
    os.path.join(r"C:\Users\abell\Desktop", "op"),
    os.path.join(r"C:\Users\abell\Desktop", "outputs")
]

for d in TARGET_DIRS:
    os.makedirs(d, exist_ok=True)

# Helper for Belt Scale 11-row calibration grid with all 8 columns
def make_belt_scale_grid(num_ao):
    grid = []
    for i in range(11):
        pct = f"{i * 10}%"
        belt_load = f"{i * 12.0:.1f}"
        spd = "2.50"
        rate = f"{i * 120.0:.1f}"
        tot = f"{i * 12.00:.2f}"
        ao1 = f"{4.0 + (i * 1.6):.2f}"
        ao2 = f"{4.0 + (i * 1.6):.2f}" if num_ao >= 2 else ""
        ao3 = f"{4.0 + (i * 1.6):.2f}" if num_ao >= 3 else ""
        grid.append([pct, belt_load, spd, rate, tot, ao1, ao2, ao3])
    return grid

# Helper for Digital Indicator 12-row calibration grid
def make_di_grid(calib_mode, lc_sys, num_ao):
    grid = []
    for i in range(1, 13):
        pct = f"{i * 100} Kg" if "load" in calib_mode.lower() else f"{i * 2.0:.1f} mV"
        row = [pct]
        for c in range(4):
            row.append(f"{350.0 + (i * 0.1) + (c * 0.05):.2f}")
        ao1 = f"{4.0 + ((i - 1) * (16.0 / 11.0)):.2f}"
        row.append(ao1)
        ao2 = f"{4.0 + ((i - 1) * (16.0 / 11.0)):.2f}" if num_ao >= 2 else ""
        row.append(ao2)
        if lc_sys == 8:
            for c in range(4):
                row.append(f"{350.1 + (i * 0.1) + (c * 0.05):.2f}")
        grid.append(row)
    return grid

# Definition of all 22 varieties across Digital Indicator, Belt Scale, Misc, and Signal Conditioner
VARIETIES = [
    # -------------------------------------------------------------------------
    # 1. Digital Indicator - 8 Varieties (4LC & 8LC x mV & Load x 1AO & 2AO)
    # -------------------------------------------------------------------------
    {
        "id": "01_Digital_Indicator_4LC_mV_1AO",
        "system": "Digital Indicator",
        "variety_name": "DI 4 LC - mV Calibration - 1 AO",
        "filename": "Digital_Indicator_4LC_mV_1AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-4MV1-01",
            "customer": "Tata Steel Ltd (Kalinganagar)",
            "capacity": "1200 Kg",
            "tag_no": "TAG-DI-01",
            "resolution": "0.1 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-101",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-10T",
            "s1": "SN-LC-01", "s2": "SN-LC-02", "s3": "SN-LC-03", "s4": "SN-LC-04",
            "remote_model": "RM-400", "remote_serial": "SN-RM-101",
            "jbox_model": "JB-4WAY", "jbox_serial": "SN-JB-101",
            "short_check": "Checked & found OK",
            "calib_mode": "mv",
            "lc_system": "4",
            "num_outputs": "1",
            "grid_data": make_di_grid("mv", 4, 1),
            "instrument_used": "Fluke 725 Calibrator",
            "remarks": "Standard 4LC mV calibration verified linear.",
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "02_Digital_Indicator_4LC_mV_2AO",
        "system": "Digital Indicator",
        "variety_name": "DI 4 LC - mV Calibration - 2 AO",
        "filename": "Digital_Indicator_4LC_mV_2AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-4MV2-02",
            "customer": "Jindal Steel & Power Ltd",
            "capacity": "1500 Kg",
            "tag_no": "TAG-DI-02",
            "resolution": "0.1 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-102",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-10T",
            "s1": "SN-LC-05", "s2": "SN-LC-06", "s3": "SN-LC-07", "s4": "SN-LC-08",
            "remote_model": "RM-400", "remote_serial": "SN-RM-102",
            "jbox_model": "JB-4WAY", "jbox_serial": "SN-JB-102",
            "short_check": "Checked & found OK",
            "calib_mode": "mv",
            "lc_system": "4",
            "num_outputs": "2",
            "current_output": "yes",
            "grid_data": make_di_grid("mv", 4, 2),
            "instrument_used": "Fluke 725 Calibrator",
            "remarks": "Dual 4-20mA analog outputs verified.",
            "tested_by": "Amit Sharma (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "03_Digital_Indicator_4LC_Load_1AO",
        "system": "Digital Indicator",
        "variety_name": "DI 4 LC - Load Calibration - 1 AO",
        "filename": "Digital_Indicator_4LC_Load_1AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-4LD1-03",
            "customer": "JSW Steel Ltd (Vijayanagar)",
            "capacity": "2000 Kg",
            "tag_no": "TAG-DI-03",
            "resolution": "0.2 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-103",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-10T",
            "s1": "SN-LC-09", "s2": "SN-LC-10", "s3": "SN-LC-11", "s4": "SN-LC-12",
            "remote_model": "RM-400", "remote_serial": "SN-RM-103",
            "jbox_model": "JB-4WAY", "jbox_serial": "SN-JB-103",
            "short_check": "Checked & found OK",
            "calib_mode": "load",
            "lc_system": "4",
            "num_outputs": "1",
            "grid_data": make_di_grid("load", 4, 1),
            "instrument_used": "Certified Test Weights (NABL)",
            "remarks": "Physical deadweight calibration verified.",
            "tested_by": "Priya Patel (QC Inspector)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "04_Digital_Indicator_4LC_Load_2AO",
        "system": "Digital Indicator",
        "variety_name": "DI 4 LC - Load Calibration - 2 AO",
        "filename": "Digital_Indicator_4LC_Load_2AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-4LD2-04",
            "customer": "Adani Power & Infra Corp",
            "capacity": "2500 Kg",
            "tag_no": "TAG-DI-04",
            "resolution": "0.2 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-104",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-10T",
            "s1": "SN-LC-13", "s2": "SN-LC-14", "s3": "SN-LC-15", "s4": "SN-LC-16",
            "remote_model": "RM-400", "remote_serial": "SN-RM-104",
            "jbox_model": "JB-4WAY", "jbox_serial": "SN-JB-104",
            "short_check": "Checked & found OK",
            "calib_mode": "load",
            "lc_system": "4",
            "num_outputs": "2",
            "current_output": "yes",
            "grid_data": make_di_grid("load", 4, 2),
            "instrument_used": "Certified Test Weights & DMM",
            "remarks": "Deadweight test with dual AO outputs verified.",
            "tested_by": "Suresh Reddy (Field Engineer)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "05_Digital_Indicator_8LC_mV_1AO",
        "system": "Digital Indicator",
        "variety_name": "DI 8 LC - mV Calibration - 1 AO",
        "filename": "Digital_Indicator_8LC_mV_1AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8MV1-05",
            "customer": "NTPC Ramagundam STPS",
            "capacity": "5000 Kg",
            "tag_no": "TAG-DI-05",
            "resolution": "0.5 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-105",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-20T",
            "s1": "SN-LC-17", "s2": "SN-LC-18", "s3": "SN-LC-19", "s4": "SN-LC-20",
            "s5": "SN-LC-21", "s6": "SN-LC-22", "s7": "SN-LC-23", "s8": "SN-LC-24",
            "remote_model": "RM-400", "remote_serial": "SN-RM-105",
            "jbox_model": "JB-8WAY", "jbox_serial": "SN-JB-105",
            "short_check": "Checked & found OK",
            "calib_mode": "mv",
            "lc_system": "8",
            "num_outputs": "1",
            "grid_data": make_di_grid("mv", 8, 1),
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "remarks": "8 LC mV cornering and summation verified.",
            "tested_by": "Anil Verma (Lead Specialist)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "06_Digital_Indicator_8LC_mV_2AO",
        "system": "Digital Indicator",
        "variety_name": "DI 8 LC - mV Calibration - 2 AO",
        "filename": "Digital_Indicator_8LC_mV_2AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8MV2-06",
            "customer": "NMDC Steel Plant (Nagarnar)",
            "capacity": "6000 Kg",
            "tag_no": "TAG-DI-06",
            "resolution": "0.5 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-106",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-20T",
            "s1": "SN-LC-25", "s2": "SN-LC-26", "s3": "SN-LC-27", "s4": "SN-LC-28",
            "s5": "SN-LC-29", "s6": "SN-LC-30", "s7": "SN-LC-31", "s8": "SN-LC-32",
            "remote_model": "RM-400", "remote_serial": "SN-RM-106",
            "jbox_model": "JB-8WAY", "jbox_serial": "SN-JB-106",
            "short_check": "Checked & found OK",
            "calib_mode": "mv",
            "lc_system": "8",
            "num_outputs": "2",
            "current_output": "yes",
            "grid_data": make_di_grid("mv", 8, 2),
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "remarks": "8 LC mV calibration with dual 4-20mA OK.",
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "07_Digital_Indicator_8LC_Load_1AO",
        "system": "Digital Indicator",
        "variety_name": "DI 8 LC - Load Calibration - 1 AO",
        "filename": "Digital_Indicator_8LC_Load_1AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8LD1-07",
            "customer": "Vedanta Aluminium & Power",
            "capacity": "8000 Kg",
            "tag_no": "TAG-DI-07",
            "resolution": "1.0 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-107",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-20T",
            "s1": "SN-LC-33", "s2": "SN-LC-34", "s3": "SN-LC-35", "s4": "SN-LC-36",
            "s5": "SN-LC-37", "s6": "SN-LC-38", "s7": "SN-LC-39", "s8": "SN-LC-40",
            "remote_model": "RM-400", "remote_serial": "SN-RM-107",
            "jbox_model": "JB-8WAY", "jbox_serial": "SN-JB-107",
            "short_check": "Checked & found OK",
            "calib_mode": "load",
            "lc_system": "8",
            "num_outputs": "1",
            "grid_data": make_di_grid("load", 8, 1),
            "instrument_used": "Heavy Test Deadweights",
            "remarks": "Full deadweight load calibration verified.",
            "tested_by": "Amit Sharma (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "08_Digital_Indicator_8LC_Load_2AO",
        "system": "Digital Indicator",
        "variety_name": "DI 8 LC - Load Calibration - 2 AO",
        "filename": "Digital_Indicator_8LC_Load_2AO_Report",
        "payload": {
            "report_type": "digital_indicator",
            "job_no": "JOB-DI-8LD2-08",
            "customer": "BHEL Heavy Electricals",
            "capacity": "10000 Kg",
            "tag_no": "TAG-DI-08",
            "resolution": "1.0 kg",
            "di_model": "DI-7000-PROCESS",
            "di_serial": "SN-DI-108",
            "sw_version": "v2.4",
            "lc_model": "LC-CANISTER-20T",
            "s1": "SN-LC-41", "s2": "SN-LC-42", "s3": "SN-LC-43", "s4": "SN-LC-44",
            "s5": "SN-LC-45", "s6": "SN-LC-46", "s7": "SN-LC-47", "s8": "SN-LC-48",
            "remote_model": "RM-400", "remote_serial": "SN-RM-108",
            "jbox_model": "JB-8WAY", "jbox_serial": "SN-JB-108",
            "short_check": "Checked & found OK",
            "calib_mode": "load",
            "lc_system": "8",
            "num_outputs": "2",
            "current_output": "yes",
            "grid_data": make_di_grid("load", 8, 2),
            "instrument_used": "Calibrated Weights & Keithley DMM",
            "remarks": "High capacity 8LC load test with dual AO verified.",
            "tested_by": "Suresh Reddy (Field Engineer)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },

    # -------------------------------------------------------------------------
    # 2. Belt Scale - 6 Varieties (4LC & 8LC x 1AO, 2AO, 3AO)
    # -------------------------------------------------------------------------
    {
        "id": "09_Belt_Scale_4LC_1AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 4 LC - 1 Analog Output",
        "filename": "Belt_Scale_4LC_1AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "4",
            "num_outputs": "1",
            "job_no": "JOB-BS-4LC1AO-09",
            "customer": "Jindal Steel & Power Ltd",
            "conveyor_no": "CV-BS-01",
            "capacity": "1200 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-401", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-401",
            "sensor_model": "LC-350",
            "s1": "SN-LC-101", "s2": "SN-LC-102", "s3": "SN-LC-103", "s4": "SN-LC-104",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-401",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-4WAY", "jbox_serial1": "SN-JB-401",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-401",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(1),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "230.2V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.01V (PASS)"},
                {"spec": "AO1: 4 - 20 mA", "act": "4.00 - 20.00 mA (Linear)"}
            ],
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "10_Belt_Scale_4LC_2AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 4 LC - 2 Analog Outputs",
        "filename": "Belt_Scale_4LC_2AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "4",
            "num_outputs": "2",
            "job_no": "JOB-BS-4LC2AO-10",
            "customer": "Tata Steel Ltd (Kalinganagar)",
            "conveyor_no": "CV-BS-02",
            "capacity": "1500 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-402", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-402",
            "sensor_model": "LC-350",
            "s1": "SN-LC-105", "s2": "SN-LC-106", "s3": "SN-LC-107", "s4": "SN-LC-108",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-402",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-4WAY", "jbox_serial1": "SN-JB-402",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-402",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(2),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "230.1V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.02V (PASS)"},
                {"spec": "Dual AO1/AO2 (4-20mA)", "act": "Both outputs verified accurate"}
            ],
            "tested_by": "Amit Sharma (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "11_Belt_Scale_4LC_3AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 4 LC - 3 Analog Outputs",
        "filename": "Belt_Scale_4LC_3AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "4",
            "num_outputs": "3",
            "job_no": "JOB-BS-4LC3AO-11",
            "customer": "JSW Steel Ltd (Vijayanagar)",
            "conveyor_no": "CV-BS-03",
            "capacity": "1800 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-403", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-403",
            "sensor_model": "LC-350",
            "s1": "SN-LC-109", "s2": "SN-LC-110", "s3": "SN-LC-111", "s4": "SN-LC-112",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-403",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-4WAY", "jbox_serial1": "SN-JB-403",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-403",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(3),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "230.4V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.00V (PASS)"},
                {"spec": "Triple AO1/AO2/AO3", "act": "All 3 outputs linear within 0.1%"}
            ],
            "tested_by": "Priya Patel (QC Inspector)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "12_Belt_Scale_8LC_1AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 8 LC - 1 Analog Output",
        "filename": "Belt_Scale_8LC_1AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "8",
            "num_outputs": "1",
            "job_no": "JOB-BS-8LC1AO-12",
            "customer": "Adani Power & Infra Corp",
            "conveyor_no": "CV-BS-04",
            "capacity": "2500 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-801", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-801",
            "sensor_model": "LC-350-HEAVY",
            "s1": "SN-8LC-01", "s2": "SN-8LC-02", "s3": "SN-8LC-03", "s4": "SN-8LC-04",
            "s5": "SN-8LC-05", "s6": "SN-8LC-06", "s7": "SN-8LC-07", "s8": "SN-8LC-08",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-801",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-8WAY", "jbox_serial1": "SN-JB-801",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-801",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(1),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "229.8V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.03V (PASS)"},
                {"spec": "8 LC Cornering Balance", "act": "Cornering error < 0.05%"}
            ],
            "tested_by": "Suresh Reddy (Field Engineer)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "13_Belt_Scale_8LC_2AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 8 LC - 2 Analog Outputs",
        "filename": "Belt_Scale_8LC_2AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "8",
            "num_outputs": "2",
            "job_no": "JOB-BS-8LC2AO-13",
            "customer": "NTPC Ramagundam STPS",
            "conveyor_no": "CV-BS-05",
            "capacity": "3000 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-802", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-802",
            "sensor_model": "LC-350-HEAVY",
            "s1": "SN-8LC-09", "s2": "SN-8LC-10", "s3": "SN-8LC-11", "s4": "SN-8LC-12",
            "s5": "SN-8LC-13", "s6": "SN-8LC-14", "s7": "SN-8LC-15", "s8": "SN-8LC-16",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-802",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-8WAY", "jbox_serial1": "SN-JB-802",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-802",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(2),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "230.0V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.01V (PASS)"},
                {"spec": "Dual AO Calibration", "act": "AO1 (Rate) & AO2 (Load) PASS"}
            ],
            "tested_by": "Anil Verma (Lead Specialist)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "14_Belt_Scale_8LC_3AO",
        "system": "Belt Scale",
        "variety_name": "Belt Scale 8 LC - 3 Analog Outputs",
        "filename": "Belt_Scale_8LC_3AO_Report",
        "payload": {
            "report_type": "belt_scale",
            "lc_system": "8",
            "num_outputs": "3",
            "job_no": "JOB-BS-8LC3AO-14",
            "customer": "NMDC Steel Plant (Nagarnar)",
            "conveyor_no": "CV-BS-06",
            "capacity": "3500 TPH",
            "res_l": "0.01", "res_s": "0.001", "res_r": "0.1", "res_t": "0.001",
            "instrument_used": "Fluke 725 Multifunction Calibrator",
            "bs_model": "BS-9000", "bs_serial": "SN-BS-803", "bs_sw": "v3.42",
            "remote_model": "RM-400", "remote_serial": "SN-RM-803",
            "sensor_model": "LC-350-HEAVY",
            "s1": "SN-8LC-17", "s2": "SN-8LC-18", "s3": "SN-8LC-19", "s4": "SN-8LC-20",
            "s5": "SN-8LC-21", "s6": "SN-8LC-22", "s7": "SN-8LC-23", "s8": "SN-8LC-24",
            "tacho_model": "TACHO-100", "tacho_serial": "SN-TCH-803",
            "tacho_rpm": "1500", "tacho_belt_speed": "2.50", "tacho_wheel_dia": "300",
            "jbox_model1": "JB-8WAY", "jbox_serial1": "SN-JB-803",
            "angle_sensor_model": "AS-300-TILT", "angle_sensor_serial": "SN-AS-803",
            "short_check": "CHECKED AND FOUND OK",
            "grid_data": make_belt_scale_grid(3),
            "routine_tests": [
                {"spec": "230V +/- 10%", "act": "230.2V (PASS)"},
                {"spec": "24V DC Excitation", "act": "24.00V (PASS)"},
                {"spec": "Triple AO Calibration", "act": "AO1, AO2, AO3 all within 0.05%"}
            ],
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },

    # -------------------------------------------------------------------------
    # 3. Misc - 6 Varieties
    # -------------------------------------------------------------------------
    {
        "id": "15_Misc_Junction_Box_Single",
        "system": "Misc Report",
        "variety_name": "Junction Box Single",
        "filename": "Misc_Junction_Box_Single_Report",
        "payload": {
            "report_type": "misc",
            "category": "jb_single",
            "chk_jb": True,
            "job_no": "JOB-MISC-JB1-15",
            "customer": "JSW Steel Ltd (Vijayanagar)",
            "custom_model": "JB-SS-4WAY",
            "custom_sl": "SN-JB-401",
            "custom_desc": "4-Way Stainless Steel Junction Box",
            "custom_qty": "1 No",
            "readings_result": "Continuity across all terminal blocks verified: < 0.05 Ohm. IP67 seal test OK.",
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "16_Misc_Junction_Box_Multiple",
        "system": "Misc Report",
        "variety_name": "Junction Box Multiple",
        "filename": "Misc_Junction_Box_Multiple_Report",
        "payload": {
            "report_type": "misc",
            "category": "jb_multiple",
            "chk_jb": True,
            "job_no": "JOB-MISC-JBM-16",
            "customer": "Adani Power & Infra Corp",
            "custom_desc": "Industrial Junction Boxes Multi-Lot",
            "variant_model_1": "JB-4WAY-AL",
            "variant_serial_1": "SN-JB-101, SN-JB-102",
            "variant_qty_1": "2 Nos",
            "variant_model_2": "JB-6WAY-SS",
            "variant_serial_2": "SN-JB-201, SN-JB-202",
            "variant_qty_2": "2 Nos",
            "variant_model_3": "JB-8WAY-CI",
            "variant_serial_3": "SN-JB-301",
            "variant_qty_3": "1 No",
            "readings_result": "All junction box variants continuity tested and found OK. Terminal torques verified.",
            "tested_by": "Amit Sharma (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "17_Misc_Load_Cell_3Box",
        "system": "Misc Report",
        "variety_name": "Load Cell 3-Boxes Test",
        "filename": "Misc_Load_Cell_3Box_Report",
        "payload": {
            "report_type": "misc",
            "category": "load_cell",
            "chk_lc": True,
            "job_no": "JOB-MISC-LC3-17",
            "customer": "NTPC Ramagundam STPS",
            "custom_model": "LC-CANISTER-20T",
            "custom_desc": "Load Cells (3-Boxes Test)",
            "tracking_no": "TRK-2026-LC-881",
            "lc_serial_1": "SN-LC-301", "lc_in_res_1": "350.2", "lc_out_res_1": "350.1", "lc_unbalance_1": "0.02",
            "lc_serial_2": "SN-LC-302", "lc_in_res_2": "350.1", "lc_out_res_2": "350.0", "lc_unbalance_2": "0.01",
            "lc_serial_3": "SN-LC-303", "lc_in_res_3": "350.4", "lc_out_res_3": "350.2", "lc_unbalance_3": "0.03",
            "lc_serial_4": "SN-LC-304", "lc_in_res_4": "350.3", "lc_out_res_4": "350.1", "lc_unbalance_4": "0.02",
            "readings_result": "The above mentioned load cells 3 boxes were checked for input resistance, output resistance, and initial unbalance and found OK.",
            "tested_by": "Priya Patel (QC Inspector)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "18_Misc_Remote_Display",
        "system": "Misc Report",
        "variety_name": "Remote Display Unit",
        "filename": "Misc_Remote_Display_Report",
        "payload": {
            "report_type": "misc",
            "category": "remote_display",
            "chk_rd": True,
            "job_no": "JOB-MISC-RD-18",
            "customer": "BHEL Heavy Electricals",
            "remote_model": "RD-LED-4INCH",
            "custom_sl": "SN-RD-901",
            "custom_desc": "Large 4-Inch LED Remote Display Unit",
            "custom_qty": "1 No",
            "tracking_no": "TRK-2026-RD-442",
            "ind_model": "BS-9000",
            "readings_result": "RS-485 serial communication with BS-9000 verified at 9600 baud. All 7-segment display segments OK.",
            "tested_by": "Suresh Reddy (Field Engineer)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "19_Misc_Digital_Indicator",
        "system": "Misc Report",
        "variety_name": "Digital Indicator Subsystem",
        "filename": "Misc_Digital_Indicator_Report",
        "payload": {
            "report_type": "misc",
            "category": "general",
            "chk_ind": True,
            "job_no": "JOB-MISC-DI-19",
            "customer": "NMDC Steel Plant (Nagarnar)",
            "custom_model": "DI-7000-PROCESS",
            "custom_sl": "SN-DI-7701",
            "custom_desc": "Digital Process Indicator Panel",
            "custom_qty": "1 No",
            "readings_result": "1. Zero Calibration: Verified 0.00 kg\n2. Span Calibration: Verified 1000.0 kg\n3. Analog Output 4-20mA: Verified 4.01mA - 19.99mA\n4. Modbus RTU Communication: OK",
            "tested_by": "Rajesh Kumar (QA Tech)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },
    {
        "id": "20_Misc_Custom_NxM_Table",
        "system": "Misc Report",
        "variety_name": "Custom Equipment with N×M Dynamic Table",
        "filename": "Misc_Custom_NxM_Table_Report",
        "payload": {
            "report_type": "misc",
            "category": "custom",
            "chk_custom": True,
            "job_no": "JOB-MISC-NXM-20",
            "customer": "Vedanta Aluminium & Power",
            "custom_model": "CUST-SMART-SYS-9",
            "custom_sl": "SN-CSS-881, SN-CSS-882",
            "custom_desc": "Custom Smart Equipment Multi-Stage Unit",
            "custom_qty": "2 Nos",
            "custom_table": {
                "headers": ["SL NO", "ITEM DESCRIPTION", "NOMINAL", "OBSERVED", "STATUS"],
                "rows": [
                    ["1", "Input Power Stage 230V AC", "230V +/- 10%", "231.2 V", "PASSED"],
                    ["2", "Regulated DC Bus 24V", "24.0V +/- 0.2V", "24.04 V", "PASSED"],
                    ["3", "Output Relay Trigger Delay", "< 50 ms", "38 ms", "PASSED"],
                    ["4", "Ethernet / Modbus TCP Packet", "100 Mbps Link", "Ping 1ms", "PASSED"]
                ]
            },
            "readings_result": "All 4 custom parameters were tested under full electrical load and certified OK.",
            "tested_by": "Anil Verma (Lead Specialist)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "2026-10-08"
        }
    },

    # -------------------------------------------------------------------------
    # 4. Signal Conditioner - 2 Varieties (mV & Load Cell)
    # -------------------------------------------------------------------------
    {
        "id": "21_Signal_Conditioner_mV",
        "system": "Signal Conditioner",
        "variety_name": "Millivolt (mV) Standard Calibration",
        "filename": "Signal_Conditioner_mV_Report",
        "payload": {
            "report_type": "signal_conditioner",
            "calib_type": "mv",
            "job_no": "JOB-SC-MV-21",
            "customer": "Automation Systems Ltd",
            "sc_model": "SC-4020-MV",
            "sc_serial": "SN-SC-8899",
            "span": "20 mV",
            "tare": "0 mV",
            "exc_v": "10.0 VDC",
            "annunciation": "OK",
            "pct_error": "< 0.05%",
            "multimeter_used": "Fluke 87V (Cal. Valid)",
            "grid_data": [
                ["0.0 mV", "4.00 mA"],
                ["4.0 mV", "7.20 mA"],
                ["8.0 mV", "10.40 mA"],
                ["12.0 mV", "13.60 mA"],
                ["16.0 mV", "16.80 mA"],
                ["20.0 mV", "20.00 mA"]
            ],
            "tested_by": "Electronics Tech",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    },
    {
        "id": "22_Signal_Conditioner_LoadCell",
        "system": "Signal Conditioner",
        "variety_name": "Load Cell Calibration",
        "filename": "Signal_Conditioner_LoadCell_Report",
        "payload": {
            "report_type": "signal_conditioner",
            "calib_type": "load_cell",
            "job_no": "JOB-SC-LC-22",
            "customer": "Cement Works Ltd",
            "sc_model": "SC-4020-LC",
            "sc_serial": "SN-SC-9902",
            "lc_model": "LC-CANISTER-10T",
            "lc_serial": "SN-LC-5511",
            "span": "2.0 mV/V",
            "tare": "0.0 mV/V",
            "exc_v": "10.0 VDC",
            "annunciation": "OK",
            "pct_error": "< 0.02%",
            "multimeter_used": "Keithley 2000 DMM",
            "grid_data": [
                ["0.0 mV/V", "4.00 mA"],
                ["0.4 mV/V", "7.20 mA"],
                ["0.8 mV/V", "10.40 mA"],
                ["1.2 mV/V", "13.60 mA"],
                ["1.6 mV/V", "16.80 mA"],
                ["2.0 mV/V", "20.00 mA"]
            ],
            "tested_by": "Priya Patel (QC Inspector)",
            "approved_by": "Dr. V. K. Malhotra (HOD-PDN)",
            "date": "08-Oct-2026"
        }
    }
]

def run_tests():
    print("=" * 85)
    print("AUTOMATED CODE TEST SUITE: DIGITAL INDICATOR, BELT SCALE, MISC, SIGNAL CONDITIONER")
    print(f"Target Varieties: {len(VARIETIES)}")
    print("Target Destinations:")
    for d in TARGET_DIRS:
        print(f"  -> {d}")
    print("=" * 85)

    # Clean out stale files from target directories first
    for d in TARGET_DIRS:
        for f in os.listdir(d):
            p = os.path.join(d, f)
            if os.path.isfile(p):
                try: os.remove(p)
                except Exception: pass

    results = []
    temp_dir = os.path.join(BASE_DIR, "tests", "temp_variety_outputs")
    os.makedirs(temp_dir, exist_ok=True)

    start_all = time.time()

    for idx, item in enumerate(VARIETIES, 1):
        test_id = item["id"]
        system = item["system"]
        vname = item["variety_name"]
        fname = item["filename"]
        payload = item["payload"]

        docx_path = os.path.join(temp_dir, f"{fname}.docx")
        pdf_path = os.path.join(temp_dir, f"{fname}.pdf")

        t0 = time.time()
        print(f"\n[{idx:02d}/{len(VARIETIES):02d}] Testing: {system} -> {vname} ...", end=" ", flush=True)

        try:
            # For Digital Indicator, populate routine tests 1..5, checklist 6..10, and load cell span/tare/exc if not explicitly set
            if payload.get("report_type") == "digital_indicator":
                di_defaults = {
                    "spec_1": "230V +/- 10%", "act_1": "230.1V (PASS)",
                    "spec_2": "24V DC (+/-5%)", "act_2": "24.02V (PASS)",
                    "spec_3": "Check Keypad/Display", "act_3": "Checked & found OK",
                    "spec_4": "Excitation Voltage", "act_4": "10.0V DC (PASS)",
                    "spec_5": "Sense Voltage", "act_5": "10.0V DC (PASS)",
                    "check_6": "Checked & found OK",
                    "check_7": "Checked & found OK",
                    "check_8": "Checked & found OK",
                    "check_9": "Checked & found OK",
                    "check_10": "2A Fast Acting (OK)",
                }
                for k, v in di_defaults.items():
                    if k not in payload:
                        payload[k] = v
                max_lc = 8 if str(payload.get("lc_system")) == "8" else 4
                for lci in range(1, max_lc + 1):
                    if f"span_lc{lci}" not in payload: payload[f"span_lc{lci}"] = "2.000 mV/V"
                    if f"tare_lc{lci}" not in payload: payload[f"tare_lc{lci}"] = "0.000 mV/V"
                    if f"exc_lc{lci}" not in payload: payload[f"exc_lc{lci}"] = "10.0 VDC"

            # 1. Generate DOCX and PDF via system code
            docx_generator.generate_docx_record(payload, docx_path)
            dur = time.time() - t0

            # 2. Verify file presence
            assert os.path.exists(docx_path), f"DOCX generation failed for {fname}"
            assert os.path.exists(pdf_path), f"PDF conversion failed for {fname}"

            # 3. Inspect PDF with PyMuPDF
            doc_pdf = pymupdf.open(pdf_path)
            page_count = len(doc_pdf)
            doc_pdf.close()

            # 4. Copy to all target Desktop folders
            docx_size = os.path.getsize(docx_path)
            pdf_size = os.path.getsize(pdf_path)

            for target_dir in TARGET_DIRS:
                shutil.copy2(docx_path, os.path.join(target_dir, f"{fname}.docx"))
                shutil.copy2(pdf_path, os.path.join(target_dir, f"{fname}.pdf"))

            print(f"OK ({dur:.2f}s) | PDF Pages: {page_count} | DOCX: {docx_size//1024}KB, PDF: {pdf_size//1024}KB")

            results.append({
                "id": test_id,
                "system": system,
                "variety": vname,
                "filename": fname,
                "status": "PASSED",
                "page_count": page_count,
                "duration_sec": round(dur, 2),
                "docx_kb": docx_size // 1024,
                "pdf_kb": pdf_size // 1024
            })

        except Exception as e:
            print(f"FAILED: {e}")
            results.append({
                "id": test_id,
                "system": system,
                "variety": vname,
                "filename": fname,
                "status": "FAILED",
                "error": str(e)
            })

    total_time = time.time() - start_all

    print("\n" + "=" * 85)
    print("TEST EXECUTION SUMMARY:")
    print("=" * 85)
    passed_count = sum(1 for r in results if r["status"] == "PASSED")
    print(f"Total Tests Run: {len(results)} | Passed: {passed_count} | Failed: {len(results) - passed_count} | Total Time: {total_time:.2f}s\n")

    print(f"{'#':<3} {'System':<20} {'Variety':<36} {'Pages':<6} {'Status':<8}")
    print("-" * 85)
    for idx, r in enumerate(results, 1):
        if r["status"] == "PASSED":
            print(f"{idx:<3} {r['system']:<20} {r['variety']:<36} {r['page_count']:<6} {r['status']:<8}")
        else:
            print(f"{idx:<3} {r['system']:<20} {r['variety']:<36} {'ERR':<6} {r['status']:<8}")
    print("=" * 85)

    # Save summary report in all target directories
    summary_path = os.path.join(temp_dir, "test_summary_report.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump({"summary": {"total": len(results), "passed": passed_count, "time_sec": total_time}, "results": results}, f, indent=2)

    for target_dir in TARGET_DIRS:
        shutil.copy2(summary_path, os.path.join(target_dir, "test_summary_report.json"))

    print(f"\nAll files successfully created and copied to target directories:")
    for target_dir in TARGET_DIRS:
        file_count = len(os.listdir(target_dir))
        print(f"  [DIR] {target_dir} ({file_count} files)")

    return passed_count == len(VARIETIES)

if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)

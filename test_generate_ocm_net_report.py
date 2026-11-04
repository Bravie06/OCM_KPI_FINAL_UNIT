import os
import tempfile
import pytest
import openpyxl
from datetime import date
from generate_ocm_net_report import (
    process_vendor_files,
    update_ocm_file,
    aggregate_weekly,
    aggregate_monthly,
    KPI_MAP,
    SHEETS,
)

def create_vendor_mock_file(path, headers, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(path)

def create_ocm_mock_file(path, sheets_cols, rows_site):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, cols in sheets_cols.items():
        ws = wb.create_sheet(title=sheet_name)
        # Row 1 empty
        ws.append([])
        # Row 2 headers: Site Name, Code du Site, Vendor, col1, col2...
        header = ['Site Name', 'Code du Site', 'Vendor'] + cols
        ws.append(header)
        for r in rows_site:
            ws.append(r)
    wb.save(path)

def test_daily_vs_weekly_monthly_traffic():
    with tempfile.TemporaryDirectory() as tmpdir:
        # H_2G (date, site, CS traffic Erl, 2G PS Traffic MB)
        h2g_path = os.path.join(tmpdir, 'H 2G test.xlsx')
        headers_2g = ['Date', 'Site Name', 'K3014:Traffic Volume on TCH(Erl)', '2G PS Traffic(MB)', 'RR307:TCH Availability(%)']
        rows_2g = [
            ['2026-05-11', 'DLA_001_A', 2000.0, 500000.0, 99.5], # 2000 Erl, 500000 MB
        ]
        create_vendor_mock_file(h2g_path, headers_2g, rows_2g)

        # H_3G: Grp_ 3G TRAFFIC – SPEECH(Erl), Grp_3G&3G+ TOTAL DATA TRAFFIC VOLUME DL&UL(GB)
        h3g_path = os.path.join(tmpdir, 'H 3G test.xlsx')
        headers_3g = ['Date', 'NODEBNAME', 'Grp_ 3G TRAFFIC – SPEECH(Erl)', 'Grp_3G&3G+ TOTAL DATA TRAFFIC VOLUME DL&UL(GB)']
        rows_3g = [
            ['2026-05-11', 'DLA_001_U', 3000.0, 50.0], # 3000 Erl, 50 GB
        ]
        create_vendor_mock_file(h3g_path, headers_3g, rows_3g)

        # N_2G: Erlang_Traffic_Carried_2G, Grp_2G_Traffic_Data_Mbytes
        n2g_path = os.path.join(tmpdir, 'N 2G test.xlsx')
        headers_n2g = ['Period start time', 'BCF name', 'Erlang_Traffic_Carried_2G', 'Grp_2G_Traffic_Data_Mbytes']
        rows_n2g = [
            ['kpi_code_1', 'kpi_code_2', 'kpi_code_3', 'kpi_code_4'], # extra header row
            ['2026-05-11', 'YAO_001_A', 1500.0, 300000.0],
        ]
        create_vendor_mock_file(n2g_path, headers_n2g, rows_n2g)

        # Z_4G: ORA_4G_Total TRAFFIC(DL+UL)(GB)
        z4g_path = os.path.join(tmpdir, 'Z 4G test.xlsx')
        headers_z4g = ['Begin Time', 'Managed Element', 'ORA_4G_Total TRAFFIC(DL+UL)(GB)']
        rows_z4g = [
            ['2026-05-11', 'BAF_001_L', 80.0],
        ]
        create_vendor_mock_file(z4g_path, headers_z4g, rows_z4g)

        # Create OCM files
        # Daily with new sheet names: 'DailyCombinedCSTrafic (Erl)', 'DailyCombinedPSTrafic (GB)'
        daily_ocm = os.path.join(tmpdir, 'ocm_daily.xlsx')
        weekly_ocm = os.path.join(tmpdir, 'ocm_weekly.xlsx')
        monthly_ocm = os.path.join(tmpdir, 'ocm_monthly.xlsx')

        sites = [
            ['DLA_001', 'DLA_001', 'HUAWEI', None],
            ['YAO_001', 'YAO_001', 'NOKIA', None],
            ['BAF_001', 'BAF_001', 'ZTE', None],
        ]

        create_ocm_mock_file(
            daily_ocm,
            {
                'Avail2G': ['2026-05-11'],
                'DailyCombinedCSTrafic (Erl)': ['2026-05-11'],
                'DailyCombinedPSTrafic (GB)': ['2026-05-11'],
            },
            sites,
        )

        create_ocm_mock_file(
            weekly_ocm,
            {
                'Avail2G': ['2026W20'],
                'DailyCombinedCSTrafic (Kerl)': ['2026W20'],
                'DailyCombinedPSTrafic (TB)': ['2026W20'],
            },
            sites,
        )

        create_ocm_mock_file(
            monthly_ocm,
            {
                'Avail2G': ['2026M05'],
                'DailyCombinedCSTrafic (Kerl)': ['2026M05'],
                'DailyCombinedPSTrafic (TB)': ['2026M05'],
            },
            sites,
        )

        vendor_files = {
            'H_2G': h2g_path,
            'H_3G': h3g_path,
            'N_2G': n2g_path,
            'Z_4G': z4g_path,
        }
        ocm_files = {
            'daily': daily_ocm,
            'weekly': weekly_ocm,
            'monthly': monthly_ocm,
        }

        process_vendor_files(vendor_files, ocm_files)

        # Check Daily file values:
        wb_daily = openpyxl.load_workbook(daily_ocm)
        ws_cs_d = wb_daily['DailyCombinedCSTrafic (Erl)']
        ws_ps_d = wb_daily['DailyCombinedPSTrafic (GB)']

        # DLA_001: row 3
        # CS: 2000 (2G) + 3000 (3G) = 5000.0 Erl
        # PS: 500000 MB * 0.001 = 500 GB + 50 GB = 550.0 GB
        assert ws_cs_d.cell(row=3, column=4).value == 5000.0
        assert ws_ps_d.cell(row=3, column=4).value == 550.0

        # YAO_001: row 4
        # CS: 1500.0 Erl
        # PS: 300000 MB * 0.001 = 300.0 GB
        assert ws_cs_d.cell(row=4, column=4).value == 1500.0
        assert ws_ps_d.cell(row=4, column=4).value == 300.0

        # BAF_001: row 5
        # PS: 80.0 GB
        assert ws_ps_d.cell(row=5, column=4).value == 80.0

        # Check Weekly file values:
        wb_weekly = openpyxl.load_workbook(weekly_ocm)
        ws_cs_w = wb_weekly['DailyCombinedCSTrafic (Kerl)']
        ws_ps_w = wb_weekly['DailyCombinedPSTrafic (TB)']

        # DLA_001:
        # CS: 5000.0 * 0.001 = 5.0 Kerl
        # PS: 0.5 + 0.05 = 0.55 TB
        assert ws_cs_w.cell(row=3, column=4).value == 5.0
        assert ws_ps_w.cell(row=3, column=4).value == 0.55

        # YAO_001:
        # CS: 1.5 Kerl
        # PS: 0.3 TB
        assert ws_cs_w.cell(row=4, column=4).value == 1.5
        assert ws_ps_w.cell(row=4, column=4).value == 0.3

        # BAF_001:
        # PS: 80 * 0.001 = 0.08 TB
        assert ws_ps_w.cell(row=5, column=4).value == 0.08

        # Check Monthly file values:
        wb_monthly = openpyxl.load_workbook(monthly_ocm)
        ws_cs_m = wb_monthly['DailyCombinedCSTrafic (Kerl)']
        ws_ps_m = wb_monthly['DailyCombinedPSTrafic (TB)']

        assert ws_cs_m.cell(row=3, column=4).value == 5.0
        assert ws_ps_m.cell(row=3, column=4).value == 0.55
        assert ws_cs_m.cell(row=4, column=4).value == 1.5
        assert ws_ps_m.cell(row=4, column=4).value == 0.3
        assert ws_ps_m.cell(row=5, column=4).value == 0.08

def test_daily_backward_compatible_sheets():
    # If Daily file still has old sheet names, it should still update with Daily units
    with tempfile.TemporaryDirectory() as tmpdir:
        h2g_path = os.path.join(tmpdir, 'H 2G test.xlsx')
        headers_2g = ['Date', 'Site Name', 'K3014:Traffic Volume on TCH(Erl)', '2G PS Traffic(MB)']
        rows_2g = [
            ['2026-05-11', 'DLA_001_A', 2000.0, 500000.0],
        ]
        create_vendor_mock_file(h2g_path, headers_2g, rows_2g)

        daily_ocm = os.path.join(tmpdir, 'ocm_daily_old_sheets.xlsx')
        sites = [['DLA_001', 'DLA_001', 'HUAWEI', None]]

        create_ocm_mock_file(
            daily_ocm,
            {
                'DailyCombinedCSTrafic (Kerl)': ['2026-05-11'],
                'DailyCombinedPSTrafic (TB)': ['2026-05-11'],
            },
            sites,
        )

        process_vendor_files({'H_2G': h2g_path}, {'daily': daily_ocm})

        wb = openpyxl.load_workbook(daily_ocm)
        assert wb['DailyCombinedCSTrafic (Kerl)'].cell(row=3, column=4).value == 2000.0
        assert wb['DailyCombinedPSTrafic (TB)'].cell(row=3, column=4).value == 500.0

def test_packet_loss_filter_mode():
    from generate_ocm_net_report import PL_SHEETS
    with tempfile.TemporaryDirectory() as tmpdir:
        # Huawei 3G packet loss
        h3g_pl_path = os.path.join(tmpdir, 'HUAWEI 3G PACKET test.xlsx')
        headers = ['Date', 'NODEBNAME', 'VS.IPPM.Forword.DropMeans(%)']
        rows = [['2026-05-11', 'DLA_001_U', 0.5]]
        create_vendor_mock_file(h3g_pl_path, headers, rows)

        daily_ocm = os.path.join(tmpdir, 'ocm_daily.xlsx')
        create_ocm_mock_file(
            daily_ocm,
            {'3G PacketLoss': ['m1', 'm2', 'm3', 'm4', 'm5', '2026-05-11']},
            [['DLA_001', 'DLA_001', 'HUAWEI', None]],
        )

        process_vendor_files({'H_3G_PL': h3g_pl_path}, {'daily': daily_ocm}, sheets_filter=PL_SHEETS)

        wb = openpyxl.load_workbook(daily_ocm)
        val = wb['3G PacketLoss'].cell(row=3, column=9).value
        assert val == 0.5

def test_all_vendors_all_traffic_kpis():
    # Test all vendor combinations for CS and PS
    # CS:
    # H_2G: K3014:Traffic Volume on TCH(Erl)
    # H_3G: Grp_ 3G TRAFFIC – SPEECH(Erl)
    # N_2G: Erlang_Traffic_Carried_2G
    # N_3G: grp_traffic_speech_erlangs
    # Z_2G: ORA_2G_CS_TRAFFIC
    # Z_3G: ORA_3G_CS Voice Traffic (Erl)
    #
    # PS:
    # H_2G: 2G PS Traffic(MB) -> MB
    # H_3G: Grp_3G&3G+ TOTAL DATA TRAFFIC VOLUME DL&UL(GB) -> GB
    # H_4G: 4G_Grp_4G/LTE TRAFFIC VOLUME(GB) -> GB
    # N_2G: Grp_2G_Traffic_Data_Mbytes -> MB
    # N_4G: Grp_4G_Traffic_New_Gbytes -> GB
    # Z_2G: ORA_2G_Traffic_Data_GB -> GB
    # Z_3G: ORA_3G_Traffic Total Data_MAC (GB) -> GB
    # Z_4G: ORA_4G_Total TRAFFIC(DL+UL)(GB) -> GB
    with tempfile.TemporaryDirectory() as tmpdir:
        # ZTE 2G
        z2g_path = os.path.join(tmpdir, 'Z Performance 2G.xlsx')
        create_vendor_mock_file(z2g_path,
            ['Begin Time', 'SITE Name', 'ORA_2G_CS_TRAFFIC', 'ORA_2G_Traffic_Data_GB'],
            [['2026-05-11', 'ZTE_001_A', 100.0, 20.0]]
        )

        # ZTE 3G
        z3g_path = os.path.join(tmpdir, 'Z 3G.xlsx')
        create_vendor_mock_file(z3g_path,
            ['Begin Time', 'NodeB Name', 'ORA_3G_CS Voice Traffic (Erl)', 'ORA_3G_Traffic Total Data_MAC (GB)'],
            [['2026-05-11', 'ZTE_001_U', 200.0, 30.0]]
        )

        # Nokia 3G
        n3g_path = os.path.join(tmpdir, 'N 3G.xlsx')
        create_vendor_mock_file(n3g_path,
            ['Period start time', 'WBTS name', 'grp_traffic_speech_erlangs'],
            [
                ['kpi1', 'kpi2', 'kpi3'],
                ['2026-05-11', 'NOK_001_U', 400.0]
            ]
        )

        # Nokia 4G
        n4g_path = os.path.join(tmpdir, 'N 4G.xlsx')
        create_vendor_mock_file(n4g_path,
            ['Period start time', 'LNBTS name', 'Grp_4G_Traffic_New_Gbytes'],
            [
                ['kpi1', 'kpi2', 'kpi3'],
                ['2026-05-11', 'NOK_001_L', 50.0]
            ]
        )

        daily_ocm = os.path.join(tmpdir, 'ocm_daily.xlsx')
        weekly_ocm = os.path.join(tmpdir, 'ocm_weekly.xlsx')

        sites = [
            ['ZTE_001', 'ZTE_001', 'ZTE', None],
            ['NOK_001', 'NOK_001', 'NOKIA', None],
        ]

        create_ocm_mock_file(
            daily_ocm,
            {
                'DailyCombinedCSTrafic (Erl)': ['2026-05-11'],
                'DailyCombinedPSTrafic (GB)': ['2026-05-11'],
            },
            sites,
        )

        create_ocm_mock_file(
            weekly_ocm,
            {
                'DailyCombinedCSTrafic (Kerl)': ['2026W20'],
                'DailyCombinedPSTrafic (TB)': ['2026W20'],
            },
            sites,
        )

        vendor_files = {
            'Z_2G': z2g_path,
            'Z_3G': z3g_path,
            'N_3G': n3g_path,
            'N_4G': n4g_path,
        }

        process_vendor_files(vendor_files, {'daily': daily_ocm, 'weekly': weekly_ocm})

        # Check ZTE:
        # CS: 100 + 200 = 300 Erl (Daily) vs 0.3 Kerl (Weekly)
        # PS: 20 + 30 = 50 GB (Daily) vs 0.05 TB (Weekly)
        wb_d = openpyxl.load_workbook(daily_ocm)
        wb_w = openpyxl.load_workbook(weekly_ocm)

        assert wb_d['DailyCombinedCSTrafic (Erl)'].cell(row=3, column=4).value == 300.0
        assert wb_d['DailyCombinedPSTrafic (GB)'].cell(row=3, column=4).value == 50.0
        assert wb_w['DailyCombinedCSTrafic (Kerl)'].cell(row=3, column=4).value == 0.3
        assert wb_w['DailyCombinedPSTrafic (TB)'].cell(row=3, column=4).value == 0.05

        # Check Nokia:
        # CS: 400 Erl (Daily) vs 0.4 Kerl (Weekly)
        # PS: 50 GB (Daily) vs 0.05 TB (Weekly)
        assert wb_d['DailyCombinedCSTrafic (Erl)'].cell(row=4, column=4).value == 400.0
        assert wb_d['DailyCombinedPSTrafic (GB)'].cell(row=4, column=4).value == 50.0
        assert wb_w['DailyCombinedCSTrafic (Kerl)'].cell(row=4, column=4).value == 0.4
        assert wb_w['DailyCombinedPSTrafic (TB)'].cell(row=4, column=4).value == 0.05

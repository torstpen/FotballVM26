import time
from pathlib import Path
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import pythoncom
import win32com.client as win32


EXCEL_VISIBLE = True

Antall_deltagere = 12


def datetime_to_days_since_1900(dt):
    """
    Konverterer Python datetime til antall dager etter
    00:00 01.01.1900.

    Klokkeslett beholdes som desimaldel av dagen.
    Eksempel: 12:00 blir .5
    """
    epoch = datetime(1900, 1, 1, 0, 0, 0)
    return (dt - epoch).total_seconds() / 86400


def open_workbook_with_retry(
    excel,
    workbook_path,
    attempts=5,
    delay_seconds=2,
    update_links=0,
    read_only=False,
    editable=True
):
    """
    Åpner en Excel-fil med flere forsøk.

    Dette hjelper hvis Excel/OneDrive/Windows fortsatt holder fila låst
    rett etter at en annen prosess har lagret/lukket den.
    """

    workbook_path = Path(workbook_path).resolve()

    if not workbook_path.exists():
        raise FileNotFoundError(f"Finner ikke Excel-filen: {workbook_path}")

    last_error = None

    for attempt in range(1, attempts + 1):
        try:
            print(
                f"   📘 Åpner workbook, forsøk {attempt}/{attempts}: "
                f"{workbook_path}"
            )

            return excel.Workbooks.Open(
                Filename=str(workbook_path),
                UpdateLinks=update_links,
                ReadOnly=read_only,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Editable=editable
            )

        except Exception as e:
            last_error = e
            print(f"   ⚠️ Klarte ikke å åpne workbook på forsøk {attempt}: {e}")

            if attempt < attempts:
                print(f"   Venter {delay_seconds} sekunder før nytt forsøk...")
                time.sleep(delay_seconds)

    raise RuntimeError(
        f"Klarte ikke å åpne Excel-filen etter {attempts} forsøk: {workbook_path}"
    ) from last_error


def normalize_excel_value(value):
    """
    Gjør verdier sammenlignbare mellom Python og Excel.
    Eksempel: 1 og 1.0 behandles likt.
    None behandles som tom tekst.
    """
    if value is None:
        return ""

    if isinstance(value, (int, float)):
        return round(float(value), 10)

    return str(value).strip()


def get_excel_row_values(ws, row_number, start_col, num_cols):
    """
    Leser én rad fra Excel og returnerer den som vanlig Python-liste.
    """
    end_col = start_col + num_cols - 1

    values = ws.Range(
        ws.Cells(row_number, start_col),
        ws.Cells(row_number, end_col)
    ).Value

    if num_cols == 1:
        return [values]

    if isinstance(values, tuple):
        if len(values) > 0 and isinstance(values[0], tuple):
            return list(values[0])

        return list(values)

    return [values]


def rows_are_identical(row_a, row_b):
    """
    Sammenligner to rader etter normalisering.
    """
    if len(row_a) != len(row_b):
        return False

    return all(
        normalize_excel_value(a) == normalize_excel_value(b)
        for a, b in zip(row_a, row_b)
    )


def export_first_chart_from_sheet(workbook, sheet_name, output_path, scale_factor=4):
    """
    Eksporterer første innebygde graf fra et gitt ark til PNG.
    Grafen skaleres midlertidig opp før eksport for bedre oppløsning.
    """

    output_path = Path(output_path).resolve()

    try:
        ws = workbook.Worksheets(sheet_name)
    except Exception:
        print(f"⚠️ Finner ikke arket '{sheet_name}', kan ikke eksportere graf.")
        return False

    try:
        chart_objects = ws.ChartObjects()
        chart_count = chart_objects.Count
    except Exception as e:
        print(f"⚠️ Klarte ikke å lese grafer fra arket '{sheet_name}'.")
        print(f"   Detalj: {e}")
        return False

    if chart_count == 0:
        print(f"⚠️ Fant ingen grafer i arket '{sheet_name}'.")
        return False

    try:
        chart_object = chart_objects.Item(1)

        # Lagre original størrelse
        original_width = chart_object.Width
        original_height = chart_object.Height

        # Gjør grafen større midlertidig
        chart_object.Width = original_width * scale_factor
        chart_object.Height = original_height * scale_factor

        chart_object.Activate()

        output_path.parent.mkdir(parents=True, exist_ok=True)

        if output_path.exists():
            try:
                output_path.unlink()
            except Exception:
                pass

        chart_object.Chart.Export(
            Filename=str(output_path),
            FilterName="PNG"
        )

        # Sett grafen tilbake til original størrelse
        chart_object.Width = original_width
        chart_object.Height = original_height

        print(f"🖼️ Graf eksportert:")
        print(f"   {output_path}")

        return True

    except Exception as e:
        print("⚠️ Klarte ikke å eksportere grafen til bilde.")
        print(f"   Detalj: {e}")
        return False


def export_poeng_chart_images(workbook, result_path, run_timestamp):
    """
    Eksporterer grafen fra arket 'poeng' til to bildefiler:

    1. poeng_graf.png
       Overskrives hver gang.

    2. poeng_graf_YYYYMMDD_HHMMSS.png
       Historisk fil med timestamp.
    """

    result_path = Path(result_path).resolve()

    image_without_timestamp = result_path.parent / "poeng_graf.png"
    image_timestamp = run_timestamp.strftime("%Y%m%d_%H%M%S")
    image_with_timestamp = result_path.parent / f"poeng_graf_{image_timestamp}.png"

    print("")
    print("🖼️ Eksporterer graf fra arket 'poeng'...")

    ok_1 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_without_timestamp
    )

    ok_2 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_with_timestamp
    )

    if ok_1 and ok_2:
        print("✅ Begge grafbildene ble lagret.")
    elif ok_1:
        print("⚠️ Graf uten timestamp ble lagret, men timestamp-versjonen feilet.")
    elif ok_2:
        print("⚠️ Graf med timestamp ble lagret, men versjonen uten timestamp feilet.")
    else:
        print("⚠️ Ingen grafbilder ble lagret.")

    print("")

    return ok_1 or ok_2
    """
    Eksporterer grafen fra arket 'poeng' til to bildefiler:

    1. poeng_graf.png
       Overskrives hver gang.

    2. poeng_graf_YYYYMMDD_HHMMSS.png
       Historisk fil med timestamp.
    """

    result_path = Path(result_path).resolve()

    image_without_timestamp = result_path.parent / "poeng_graf.png"

    image_timestamp = run_timestamp.strftime("%Y%m%d_%H%M%S")
    image_with_timestamp = result_path.parent / f"poeng_graf_{image_timestamp}.png"

    print("")
    print("🖼️ Eksporterer graf fra arket 'poeng'...")

    ok_1 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_without_timestamp
    )

    ok_2 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_with_timestamp
    )

    if ok_1 and ok_2:
        print("✅ Begge grafbildene ble lagret.")
    elif ok_1:
        print("⚠️ Graf uten timestamp ble lagret, men timestamp-versjonen feilet.")
    elif ok_2:
        print("⚠️ Graf med timestamp ble lagret, men versjonen uten timestamp feilet.")
    else:
        print("⚠️ Ingen grafbilder ble lagret.")

    print("")

    return ok_1 or ok_2
    """
    Eksporterer grafen fra arket 'poeng' til to bildefiler:

    1. poeng_graf.png
       Overskrives hver gang.

    2. poeng_graf_YYYYMMDD_HHMMSS.png
       Historisk fil med timestamp.
    """

    result_path = Path(result_path).resolve()

    image_without_timestamp = result_path.parent / "poeng_graf.png"

    image_timestamp = run_timestamp.strftime("%Y%m%d_%H%M%S")
    image_with_timestamp = result_path.parent / f"poeng_graf_{image_timestamp}.png"

    print("")
    print("🖼️ Eksporterer graf fra arket 'poeng'...")

    ok_1 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_without_timestamp
    )

    ok_2 = export_first_chart_from_sheet(
        workbook=workbook,
        sheet_name="poeng",
        output_path=image_with_timestamp
    )

    if ok_1 and ok_2:
        print("✅ Begge grafbildene ble lagret.")
    elif ok_1:
        print("⚠️ Graf uten timestamp ble lagret, men timestamp-versjonen feilet.")
    elif ok_2:
        print("⚠️ Graf med timestamp ble lagret, men versjonen uten timestamp feilet.")
    else:
        print("⚠️ Ingen grafbilder ble lagret.")

    print("")

    return ok_1 or ok_2


def update_worldcup_home_goals(excel, admin_wb, admin_file_path):
    """
    Leser hjemmemål og bortemål fra vm_2026_resultater.xlsx, ark 'Kamper',
    og skriver verdiene inn i ADMINExcelMundial202625.xlsx,
    ark 'WORLDCUP'.

    Hjemmemål skrives til kolonne AC.
    Bortemål skrives til kolonne AD.

    Radnummeret i WORLDCUP hentes fra kolonne G i Kamper.
    """

    admin_folder = Path(admin_file_path).parent
    results_file = admin_folder / "vm_2026_resultater.xlsx"
    results_path = Path(results_file).resolve()

    if not results_path.exists():
        print(f"⚠️ Finner ikke vm_2026_resultater.xlsx: {results_path}")
        return False

    print("")
    print("⚽ Oppdaterer hjemmemål og bortemål i WORLDCUP...")
    print(f"   Leser fra: {results_path}")

    results_wb = None

    try:
        results_wb = open_workbook_with_retry(
            excel=excel,
            workbook_path=results_path,
            attempts=6,
            delay_seconds=2,
            update_links=0,
            read_only=True,
            editable=False
        )

        try:
            kamper_ws = results_wb.Worksheets("Kamper")
        except Exception:
            print("❌ Finner ikke arket 'Kamper' i vm_2026_resultater.xlsx.")
            return False

        try:
            worldcup_ws = admin_wb.Worksheets("WORLDCUP")
        except Exception:
            print("❌ Finner ikke arket 'WORLDCUP' i ADMINExcelMundial202625.xlsx.")
            return False

        used_range = kamper_ws.UsedRange
        first_row = used_range.Row
        first_col = used_range.Column
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        header_row = first_row
        home_goals_col = None
        away_goals_col = None

        # Finn kolonnene for hjemmemål og bortemål basert på overskrifter
        for col in range(first_col, first_col + col_count):
            header_value = kamper_ws.Cells(header_row, col).Value

            if header_value is None:
                continue

            header_text = str(header_value).strip().lower()

            if (
                "hjemmemål" in header_text
                or "hjemme mål" in header_text
                or "hjemmemal" in header_text
                or "home goals" in header_text
                or "homegoals" in header_text
            ):
                home_goals_col = col

            if (
                "bortemål" in header_text
                or "borte mål" in header_text
                or "bortemal" in header_text
                or "away goals" in header_text
                or "awaygoals" in header_text
            ):
                away_goals_col = col

        if home_goals_col is None:
            print("❌ Fant ikke hjemmemålkolonnen i arket 'Kamper'.")
            print("   Sjekk at overskriften inneholder for eksempel 'hjemmemål'.")
            return False

        if away_goals_col is None:
            print("❌ Fant ikke bortemålkolonnen i arket 'Kamper'.")
            print("   Sjekk at overskriften inneholder for eksempel 'bortemål'.")
            return False

        target_row_col = 7          # Kolonne G i Kamper
        home_col_worldcup = 29      # Kolonne AC i WORLDCUP
        away_col_worldcup = 30      # Kolonne AD i WORLDCUP

        last_row = first_row + row_count - 1
        updates = 0

        print(f"   Hjemmemålkolonne funnet i Kamper-kolonne {home_goals_col}.")
        print(f"   Bortemålkolonne funnet i Kamper-kolonne {away_goals_col}.")
        print("   Skriver hjemmemål til WORLDCUP kolonne AC.")
        print("   Skriver bortemål til WORLDCUP kolonne AD.")

        for row in range(header_row + 1, last_row + 1):
            home_goals = kamper_ws.Cells(row, home_goals_col).Value
            away_goals = kamper_ws.Cells(row, away_goals_col).Value
            target_row = kamper_ws.Cells(row, target_row_col).Value

            # Hopp over rader uten radnummer
            if target_row in (None, ""):
                continue

            try:
                target_row = int(target_row)
            except Exception:
                print(
                    f"   Hopper over rad {row}: ugyldig radnummer "
                    f"i kolonne G: {target_row}"
                )
                continue

            wrote_something = False

            # Skriv hjemmemål til AC
            if home_goals not in (None, ""):
                worldcup_ws.Cells(target_row, home_col_worldcup).Value = home_goals
                wrote_something = True

            # Skriv bortemål til AD
            if away_goals not in (None, ""):
                worldcup_ws.Cells(target_row, away_col_worldcup).Value = away_goals
                wrote_something = True

            if wrote_something:
                updates += 1
                print(
                    f"   Kamper rad {row}: "
                    f"hjemmemål {home_goals} skrevet til WORLDCUP!AC{target_row}, "
                    f"bortemål {away_goals} skrevet til WORLDCUP!AD{target_row}"
                )

        print(f"✅ Ferdig. Oppdaterte {updates} kamper i WORLDCUP.")
        return True

    except Exception as e:
        print(f"❌ Feil ved oppdatering av mål i WORLDCUP: {e}")
        return False

    finally:
        try:
            if results_wb is not None:
                results_wb.Close(SaveChanges=False)
        except Exception:
            pass


def refresh_excel_file(file_path, run_timestamp):
    """
    Åpner Excel-filen i ekte Excel, oppdaterer eksterne koblinger,
    beregner hele arbeidsboken, lagrer og lukker.

    Leser CLAS!C5:D basert på Antall_deltagere,
    sorterer alfabetisk etter kolonne C,
    kopierer raden over første ledige rad i kolonne B i arket 'poeng',
    og skriver poengene fra kolonne D på raden under.

    Hvis nye poeng er identiske med nederste eksisterende poengrad,
    skrives det ikke inn nye verdier.

    Etterpå eksporteres grafen i Poeng-arket til:
    - poeng_graf.png
    - poeng_graf_YYYYMMDD_HHMMSS.png
    """

    full_path = Path(file_path).resolve()

    if not full_path.exists():
        print(f"❌ Finner ikke Excel-filen: {full_path}")
        return False

    result_file = full_path.parent / "vm_2026_resultater.xlsx"
    result_path = result_file.resolve()

    print("")
    print("📘 Åpner og oppdaterer Excel...")
    print(f"   Kildefil: {full_path}")
    print(f"   Resultatfil: {result_path}")

    pythoncom.CoInitialize()

    excel = None
    wb = None
    result_wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")

        excel.Visible = EXCEL_VISIBLE
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception as e:
            print("⚠️ Klarte ikke å sette Excel Calculation=Automatic.")
            print(f"   Fortsetter likevel. Detalj: {e}")

        print("   Åpner workbook og oppdaterer eksterne koblinger automatisk...")

        wb = open_workbook_with_retry(
            excel=excel,
            workbook_path=full_path,
            attempts=5,
            delay_seconds=2,
            update_links=3,
            read_only=False,
            editable=True
        )

        # -------------------------------------------------
        # OPPDATER WORLDCUP MED HJEMMEMÅL/BORTEMÅL FRA KAMPER
        # -------------------------------------------------

        poeng_updated = update_worldcup_home_goals(excel, wb, full_path)

        try:
            wb.UpdateLinks = 3
        except Exception:
            pass

        print("   Oppdaterer linker eksplisitt...")

        try:
            links = wb.LinkSources(Type=1)  # xlExcelLinks = 1

            if links:
                for link in links:
                    try:
                        link_text = str(link)

                        if "vm_2026_resultater.xlsx" in link_text.lower():
                            print(f"   Hopper over link til resultatfil: {link}")
                            continue

                        print(f"   Oppdaterer link: {link}")
                        wb.UpdateLink(Name=link, Type=1)

                    except Exception as e:
                        print(f"   Klarte ikke å oppdatere link {link}: {e}")
        except Exception:
            pass

        print("   Oppdaterer datatilkoblinger / spørringer / pivoter...")

        try:
            for connection in wb.Connections:
                try:
                    connection.OLEDBConnection.BackgroundQuery = False
                except Exception:
                    pass

                try:
                    connection.ODBCConnection.BackgroundQuery = False
                except Exception:
                    pass
        except Exception:
            pass

        wb.RefreshAll()

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        start_time = time.time()
        timeout_seconds = 180

        while time.time() - start_time < timeout_seconds:
            refreshing = False

            try:
                for connection in wb.Connections:
                    try:
                        if connection.OLEDBConnection.Refreshing:
                            refreshing = True
                            break
                    except Exception:
                        pass

                    try:
                        if connection.ODBCConnection.Refreshing:
                            refreshing = True
                            break
                    except Exception:
                        pass
            except Exception:
                pass

            if not refreshing:
                break

            time.sleep(1)

        print("   Beregner hele arbeidsboken...")

        try:
            excel.CalculateFullRebuild()
        except Exception as e:
            print("⚠️ CalculateFullRebuild feilet. Prøver CalculateFull...")
            print(f"   Detalj: {e}")

            try:
                excel.CalculateFull()
            except Exception as e:
                print("⚠️ CalculateFull feilet. Prøver vanlig Calculate...")
                print(f"   Detalj: {e}")

                try:
                    excel.Calculate()
                except Exception as e:
                    print("⚠️ Vanlig Calculate feilet også.")
                    print(f"   Detalj: {e}")

        print("   Lagrer kildefilen...")
        wb.Save()

        # -------------------------------------------------
        # LES CLAS!C5:D BASERT PÅ ANTALL DELTAGERE
        # -------------------------------------------------

        end_row = 4 + Antall_deltagere
        source_range = f"C5:D{end_row}"

        print(f"   Leser verdier fra CLAS!{source_range}...")

        try:
            source_ws = wb.Worksheets("CLAS")
        except Exception:
            print("❌ Finner ikke arket 'CLAS' i kildefilen.")
            return False

        values = source_ws.Range(source_range).Value

        print("   Verdier lest:")
        for row in values:
            print(f"   {row}")

        # -------------------------------------------------
        # SORTER ETTER KOLONNE C
        # -------------------------------------------------

        print("   Sorterer verdiene alfabetisk etter kolonne C...")

        rows = [list(row) for row in values]

        rows_sorted = sorted(
            rows,
            key=lambda row: "" if row[0] is None else str(row[0]).casefold()
        )

        print("   Sorterte verdier:")
        for row in rows_sorted:
            print(f"   {row}")

        points_row = [row[1] for row in rows_sorted]

        print("   Poeng etter sortering:")
        print(f"   {points_row}")

        # -------------------------------------------------
        # ÅPNE / OPPRETT RESULTATFIL
        # -------------------------------------------------

        print("   Åpner resultatfil...")

        if result_path.exists():
            result_wb = open_workbook_with_retry(
                excel=excel,
                workbook_path=result_path,
                attempts=6,
                delay_seconds=2,
                update_links=0,
                read_only=False,
                editable=True
            )

            if result_wb.ReadOnly:
                print("❌ Resultatfilen ble åpnet som skrivebeskyttet.")
                print("   Lukk vm_2026_resultater.xlsx hvis den er åpen i Excel.")
                return False

        else:
            print("   Resultatfil finnes ikke. Oppretter ny fil...")
            result_wb = excel.Workbooks.Add()

        # -------------------------------------------------
        # FINN ELLER OPPRETT ARKET 'poeng'
        # -------------------------------------------------

        target_ws = None

        for sheet in result_wb.Worksheets:
            if sheet.Name.lower() == "poeng":
                target_ws = sheet
                break

        if target_ws is None:
            print("   Arket 'poeng' finnes ikke. Oppretter det...")
            target_ws = result_wb.Worksheets.Add()
            target_ws.Name = "poeng"

        # -------------------------------------------------
        # SKRIV TIL ARKET 'poeng'
        # -------------------------------------------------

        print("   Finner første ledige rad i kolonne B i arket 'poeng'...")

        xlUp = -4162
        last_row_b = target_ws.Cells(target_ws.Rows.Count, 2).End(xlUp).Row

        start_col = 2  # Kolonne B
        num_cols = len(points_row)
        end_col = start_col + num_cols - 1

        # -------------------------------------------------
        # SJEKK OM NEDERSTE RAD ALLEREDE HAR SAMME POENG
        # -------------------------------------------------

        bottom_cell_value = target_ws.Cells(last_row_b, start_col).Value

        if bottom_cell_value not in (None, ""):
            existing_bottom_points = get_excel_row_values(
                ws=target_ws,
                row_number=last_row_b,
                start_col=start_col,
                num_cols=num_cols
            )

            print("   Nederste eksisterende poengrad:")
            print(f"   {existing_bottom_points}")

            print("   Nye poeng som vurderes skrevet:")
            print(f"   {points_row}")

            if rows_are_identical(points_row, existing_bottom_points):
                print("")
                print("ℹ️ Poengene er identiske med nederste rad i arket 'poeng'.")
                print("   Skriver derfor ikke inn nye verdier.")
                print("")

                print("   Poeng ble ikke endret, hopper over grafeksport.")

                result_wb.Close(SaveChanges=False)
                result_wb = None

                print("   Lukker kildefilen...")
                wb.Close(SaveChanges=True)
                wb = None

                print("✅ Ferdig. Ingen nye poeng ble skrevet, og grafbilder ble ikke oppdatert.")
                print("")
                return True

        if target_ws.Cells(last_row_b, 2).Value in (None, ""):
            start_row = last_row_b
        else:
            start_row = last_row_b + 1

        points_row_number = start_row

        print(f"   Første ledige rad i kolonne B er rad {start_row}")

        # -------------------------------------------------
        # SKRIV TIDSPUNKT I KOLONNE A
        # -------------------------------------------------

        timestamp = datetime_to_days_since_1900(run_timestamp)

        timestamp_cell = target_ws.Cells(points_row_number, 1)
        timestamp_cell.Value2 = timestamp
        timestamp_cell.NumberFormatLocal = "0,00000000"

        print(f"   Tidspunkt skrevet til Poeng!A{points_row_number}")
        print(f"   Tidspunkt dato/klokke: {run_timestamp.strftime('%d.%m.%Y %H:%M:%S')}")
        print(f"   Tidspunkt som dager etter 01.01.1900 00:00: {timestamp}")

        # -------------------------------------------------
        # SKRIV POENGENE PÅ RADEN UNDER
        # -------------------------------------------------

        print(
            f"   Skriver poeng til "
            f"Poeng!B{points_row_number}:{target_ws.Cells(points_row_number, end_col).Address}"
        )

        target_points_range = target_ws.Range(
            target_ws.Cells(points_row_number, start_col),
            target_ws.Cells(points_row_number, end_col)
        )

        target_points_range.Value = tuple([tuple(points_row)])

        print(
            f"   Poeng skrevet til "
            f"Poeng!B{points_row_number}:{target_ws.Cells(points_row_number, end_col).Address}"
        )

        # -------------------------------------------------
        # LAGRE RESULTATFIL
        # -------------------------------------------------

        print("   Lagrer resultatfil...")

        if result_path.exists():
            result_wb.Save()
        else:
            result_wb.SaveAs(str(result_path), FileFormat=51)  # xlsx

        # -------------------------------------------------
        # EKSPORTER GRAF FRA ARKET 'poeng' TIL BILDER
        # -------------------------------------------------

        if poeng_updated:
            print("   Oppdaterer beregninger før graf eksporteres...")

            try:
                excel.CalculateFull()
            except Exception:
                try:
                    excel.Calculate()
                except Exception:
                    pass

            export_poeng_chart_images(
                workbook=result_wb,
                result_path=result_path,
                run_timestamp=run_timestamp
            )
        else:
            print("   Poeng ble ikke endret, hopper over grafeksport.")

        # Lagre igjen etter eventuell grafoppdatering
        result_wb.Save()

        result_wb.Close(SaveChanges=True)
        result_wb = None

        print("   Lukker kildefilen...")
        wb.Close(SaveChanges=True)
        wb = None

        print("✅ Excel er oppdatert, poeng er lagt til, og grafbilder er eksportert hvis nødvendig.")
        print("")

        return True

    except Exception as e:
        print("❌ Feil ved Excel-oppdatering/kopiering:", e)
        return False

    finally:
        try:
            if result_wb is not None:
                result_wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if excel is not None:
                excel.DisplayAlerts = True
                excel.EnableEvents = True
                excel.Quit()
        except Exception:
            pass

        pythoncom.CoUninitialize()
    """
    Åpner Excel-filen i ekte Excel, oppdaterer eksterne koblinger,
    beregner hele arbeidsboken, lagrer og lukker.

    Leser CLAS!C5:D basert på Antall_deltagere,
    sorterer alfabetisk etter kolonne C,
    kopierer raden over første ledige rad i kolonne B i arket 'poeng',
    og skriver poengene fra kolonne D på raden under.

    Hvis nye poeng er identiske med nederste eksisterende poengrad,
    skrives det ikke inn nye verdier.

    Etterpå eksporteres grafen i Poeng-arket til:
    - poeng_graf.png
    - poeng_graf_YYYYMMDD_HHMMSS.png
    """

    full_path = Path(file_path).resolve()

    if not full_path.exists():
        print(f"❌ Finner ikke Excel-filen: {full_path}")
        return False

    result_file = full_path.parent / "vm_2026_resultater.xlsx"
    result_path = result_file.resolve()

    print("")
    print("📘 Åpner og oppdaterer Excel...")
    print(f"   Kildefil: {full_path}")
    print(f"   Resultatfil: {result_path}")

    pythoncom.CoInitialize()

    excel = None
    wb = None
    result_wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")

        excel.Visible = EXCEL_VISIBLE
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        try:
            excel.Calculation = -4105  # xlCalculationAutomatic
        except Exception as e:
            print("⚠️ Klarte ikke å sette Excel Calculation=Automatic.")
            print(f"   Fortsetter likevel. Detalj: {e}")

        print("   Åpner workbook og oppdaterer eksterne koblinger automatisk...")

        wb = open_workbook_with_retry(
            excel=excel,
            workbook_path=full_path,
            attempts=5,
            delay_seconds=2,
            update_links=3,
            read_only=False,
            editable=True
        )

        # -------------------------------------------------
        # OPPDATER WORLDCUP MED HJEMMEMÅL/BORTEMÅL FRA KAMPER
        # -------------------------------------------------

        update_worldcup_home_goals(excel, wb, full_path)

        try:
            wb.UpdateLinks = 3
        except Exception:
            pass

        print("   Oppdaterer linker eksplisitt...")

        try:
            links = wb.LinkSources(Type=1)  # xlExcelLinks = 1

            if links:
                for link in links:
                    try:
                        link_text = str(link)

                        if "vm_2026_resultater.xlsx" in link_text.lower():
                            print(f"   Hopper over link til resultatfil: {link}")
                            continue

                        print(f"   Oppdaterer link: {link}")
                        wb.UpdateLink(Name=link, Type=1)

                    except Exception as e:
                        print(f"   Klarte ikke å oppdatere link {link}: {e}")
        except Exception:
            pass

        print("   Oppdaterer datatilkoblinger / spørringer / pivoter...")

        try:
            for connection in wb.Connections:
                try:
                    connection.OLEDBConnection.BackgroundQuery = False
                except Exception:
                    pass

                try:
                    connection.ODBCConnection.BackgroundQuery = False
                except Exception:
                    pass
        except Exception:
            pass

        wb.RefreshAll()

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        start_time = time.time()
        timeout_seconds = 180

        while time.time() - start_time < timeout_seconds:
            refreshing = False

            try:
                for connection in wb.Connections:
                    try:
                        if connection.OLEDBConnection.Refreshing:
                            refreshing = True
                            break
                    except Exception:
                        pass

                    try:
                        if connection.ODBCConnection.Refreshing:
                            refreshing = True
                            break
                    except Exception:
                        pass
            except Exception:
                pass

            if not refreshing:
                break

            time.sleep(1)

        print("   Beregner hele arbeidsboken...")

        try:
            excel.CalculateFullRebuild()
        except Exception as e:
            print("⚠️ CalculateFullRebuild feilet. Prøver CalculateFull...")
            print(f"   Detalj: {e}")

            try:
                excel.CalculateFull()
            except Exception as e:
                print("⚠️ CalculateFull feilet. Prøver vanlig Calculate...")
                print(f"   Detalj: {e}")

                try:
                    excel.Calculate()
                except Exception as e:
                    print("⚠️ Vanlig Calculate feilet også.")
                    print(f"   Detalj: {e}")

        print("   Lagrer kildefilen...")
        wb.Save()

        # -------------------------------------------------
        # LES CLAS!C5:D BASERT PÅ ANTALL DELTAGERE
        # -------------------------------------------------

        end_row = 4 + Antall_deltagere
        source_range = f"C5:D{end_row}"

        print(f"   Leser verdier fra CLAS!{source_range}...")

        try:
            source_ws = wb.Worksheets("CLAS")
        except Exception:
            print("❌ Finner ikke arket 'CLAS' i kildefilen.")
            return False

        values = source_ws.Range(source_range).Value

        print("   Verdier lest:")
        for row in values:
            print(f"   {row}")

        # -------------------------------------------------
        # SORTER ETTER KOLONNE C
        # -------------------------------------------------

        print("   Sorterer verdiene alfabetisk etter kolonne C...")

        rows = [list(row) for row in values]

        rows_sorted = sorted(
            rows,
            key=lambda row: "" if row[0] is None else str(row[0]).casefold()
        )

        print("   Sorterte verdier:")
        for row in rows_sorted:
            print(f"   {row}")

        points_row = [row[1] for row in rows_sorted]

        print("   Poeng etter sortering:")
        print(f"   {points_row}")

        # -------------------------------------------------
        # ÅPNE / OPPRETT RESULTATFIL
        # -------------------------------------------------

        print("   Åpner resultatfil...")

        if result_path.exists():
            result_wb = open_workbook_with_retry(
                excel=excel,
                workbook_path=result_path,
                attempts=6,
                delay_seconds=2,
                update_links=0,
                read_only=False,
                editable=True
            )

            if result_wb.ReadOnly:
                print("❌ Resultatfilen ble åpnet som skrivebeskyttet.")
                print("   Lukk vm_2026_resultater.xlsx hvis den er åpen i Excel.")
                return False

        else:
            print("   Resultatfil finnes ikke. Oppretter ny fil...")
            result_wb = excel.Workbooks.Add()

        # -------------------------------------------------
        # FINN ELLER OPPRETT ARKET 'poeng'
        # -------------------------------------------------

        target_ws = None

        for sheet in result_wb.Worksheets:
            if sheet.Name.lower() == "poeng":
                target_ws = sheet
                break

        if target_ws is None:
            print("   Arket 'poeng' finnes ikke. Oppretter det...")
            target_ws = result_wb.Worksheets.Add()
            target_ws.Name = "poeng"

        # -------------------------------------------------
        # SKRIV TIL ARKET 'poeng'
        # -------------------------------------------------

        print("   Finner første ledige rad i kolonne B i arket 'poeng'...")

        xlUp = -4162
        last_row_b = target_ws.Cells(target_ws.Rows.Count, 2).End(xlUp).Row

        start_col = 2  # Kolonne B
        num_cols = len(points_row)
        end_col = start_col + num_cols - 1

        # -------------------------------------------------
        # SJEKK OM NEDERSTE RAD ALLEREDE HAR SAMME POENG
        # -------------------------------------------------

        bottom_cell_value = target_ws.Cells(last_row_b, start_col).Value

        if bottom_cell_value not in (None, ""):
            existing_bottom_points = get_excel_row_values(
                ws=target_ws,
                row_number=last_row_b,
                start_col=start_col,
                num_cols=num_cols
            )

            print("   Nederste eksisterende poengrad:")
            print(f"   {existing_bottom_points}")

            print("   Nye poeng som vurderes skrevet:")
            print(f"   {points_row}")

            if rows_are_identical(points_row, existing_bottom_points):
                print("")
                print("ℹ️ Poengene er identiske med nederste rad i arket 'poeng'.")
                print("   Skriver derfor ikke inn nye verdier.")
                print("")

                print("   Oppdaterer beregninger før graf eksporteres...")

                try:
                    excel.CalculateFull()
                except Exception:
                    try:
                        excel.Calculate()
                    except Exception:
                        pass

                export_poeng_chart_images(
                    workbook=result_wb,
                    result_path=result_path,
                    run_timestamp=run_timestamp
                )

                result_wb.Close(SaveChanges=False)
                result_wb = None

                print("   Lukker kildefilen...")
                wb.Close(SaveChanges=True)
                wb = None

                print("✅ Ferdig. Ingen nye poeng ble skrevet, men grafbilder ble eksportert.")
                print("")

                return True

        if target_ws.Cells(last_row_b, 2).Value in (None, ""):
            start_row = last_row_b
        else:
            start_row = last_row_b + 1

        points_row_number = start_row

        print(f"   Første ledige rad i kolonne B er rad {start_row}")

        # -------------------------------------------------
        # SKRIV TIDSPUNKT I KOLONNE A
        # -------------------------------------------------

        timestamp = datetime_to_days_since_1900(run_timestamp)

        timestamp_cell = target_ws.Cells(points_row_number, 1)
        timestamp_cell.Value2 = timestamp
        timestamp_cell.NumberFormatLocal = "0,00000000"

        print(f"   Tidspunkt skrevet til Poeng!A{points_row_number}")
        print(f"   Tidspunkt dato/klokke: {run_timestamp.strftime('%d.%m.%Y %H:%M:%S')}")
        print(f"   Tidspunkt som dager etter 01.01.1900 00:00: {timestamp}")

        # -------------------------------------------------
        # SKRIV POENGENE PÅ RADEN UNDER
        # -------------------------------------------------

        print(
            f"   Skriver poeng til "
            f"Poeng!B{points_row_number}:{target_ws.Cells(points_row_number, end_col).Address}"
        )

        target_points_range = target_ws.Range(
            target_ws.Cells(points_row_number, start_col),
            target_ws.Cells(points_row_number, end_col)
        )

        target_points_range.Value = tuple([tuple(points_row)])

        print(
            f"   Poeng skrevet til "
            f"Poeng!B{points_row_number}:{target_ws.Cells(points_row_number, end_col).Address}"
        )

        # -------------------------------------------------
        # LAGRE RESULTATFIL
        # -------------------------------------------------

        print("   Lagrer resultatfil...")

        if result_path.exists():
            result_wb.Save()
        else:
            result_wb.SaveAs(str(result_path), FileFormat=51)  # xlsx

        # -------------------------------------------------
        # EKSPORTER GRAF FRA ARKET 'poeng' TIL BILDER
        # -------------------------------------------------

        print("   Oppdaterer beregninger før graf eksporteres...")

        try:
            excel.CalculateFull()
        except Exception:
            try:
                excel.Calculate()
            except Exception:
                pass

        export_poeng_chart_images(
            workbook=result_wb,
            result_path=result_path,
            run_timestamp=run_timestamp
        )

        # Lagre igjen etter eventuell grafoppdatering
        result_wb.Save()

        result_wb.Close(SaveChanges=True)
        result_wb = None

        print("   Lukker kildefilen...")
        wb.Close(SaveChanges=True)
        wb = None

        print("✅ Excel er oppdatert, poeng er lagt til, og grafbilder er eksportert.")
        print("")

        return True

    except Exception as e:
        print("❌ Feil ved Excel-oppdatering/kopiering:", e)
        return False

    finally:
        try:
            if result_wb is not None:
                result_wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if excel is not None:
                excel.DisplayAlerts = True
                excel.EnableEvents = True
                excel.Quit()
        except Exception:
            pass

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    run_timestamp = (
        datetime.now(ZoneInfo("Europe/Oslo")).replace(tzinfo=None)
        + timedelta(days=2)
    )

    print("🚀 Programmet starter...")
    print(f"🕒 Kjøretidspunkt: {run_timestamp.strftime('%d.%m.%Y %H:%M:%S')}")
    print(f"🕒 Dager etter 01.01.1900 00:00: {datetime_to_days_since_1900(run_timestamp)}")
    print(f"👥 Antall deltagere: {Antall_deltagere}")

    SCRIPT_DIR = Path(__file__).resolve().parent
    excel_file = SCRIPT_DIR / "ADMINExcelMundial202625.xlsx"

    print(f"📘 Skal oppdatere fil: {excel_file}")

    refresh_excel_file(excel_file, run_timestamp)

    print("✅ Programmet er ferdig.")

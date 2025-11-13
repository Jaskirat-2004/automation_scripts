import win32com.client as win32
import os
import time
from PIL import ImageGrab
from datetime import datetime
import traceback


def take_excel_snapshots(file_path, sheet_ranges, output_dir):
    """
    Take snapshots from multiple Excel sheets with different ranges.

    Args:
        file_path (str): Path to Excel file
        sheet_ranges (dict): {sheet_name: cell_range}
        output_dir (str): Directory where snapshots will be saved

    Returns:
        dict: {sheet_name: image_path}
    """

    # Create an isolated Excel instance
    excel = win32.DispatchEx("Excel.Application")
    snapshots = {}

    try:
        excel.Visible = False
        excel.DisplayAlerts = False

        # Open workbook and allow Excel to settle
        workbook = excel.Workbooks.Open(os.path.abspath(file_path))
        time.sleep(2)

        for sheet_name, cell_range in sheet_ranges.items():
            try:
                sheet = workbook.Sheets[sheet_name]
                time.sleep(1)

                # --- Retry logic for CopyPicture ---
                copy_success = False
                for attempt in range(3):
                    try:
                        sheet.Range(cell_range).CopyPicture(Format=2)
                        copy_success = True
                        break
                    except Exception as copy_error:
                        print(f"⚠️ Retry {attempt+1} for {sheet_name} due to: {copy_error}")
                        time.sleep(1)

                if not copy_success:
                    raise RuntimeError(f"❌ Failed to copy picture for {sheet_name}")

                # --- Retry grabbing clipboard image ---
                img = None
                for _ in range(5):
                    img = ImageGrab.grabclipboard()
                    if img is not None:
                        break
                    time.sleep(0.5)

                if img is None:
                    raise RuntimeError(f"❌ Failed to grab image for {sheet_name}")

                os.makedirs(output_dir, exist_ok=True)
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                image_path = os.path.join(output_dir, f"{sheet_name.replace(' ', '_')}_{timestamp}.png")

                img.save(image_path, "PNG")
                snapshots[sheet_name] = image_path
                print(f"✅ Snapshot saved for sheet '{sheet_name}' → {image_path}")

            except Exception as e:
                print(f"❌ Error capturing sheet {sheet_name}: {e}")

    except Exception as e:
        print(f"❌ Error in Excel snapshot process: {e}")
        raise

    finally:
        # Clean up Excel COM objects
        try:
            if 'workbook' in locals():
                workbook.Close(False)
            excel.Quit()
            del workbook
            del excel
        except Exception:
            pass

    # Summary output
    if len(snapshots) == 0:
        print("⚠️ No snapshots captured!")
    elif len(snapshots) == 1:
        print("📸 Single snapshot captured successfully.")
    else:
        print(f"📸 {len(snapshots)} snapshots captured successfully.")

    return snapshots


def send_snapshots_inline(
    to_emails,
    subject,
    snapshots,
    report_name,
    cc_emails=None,
    download_link=None,
    attachments=None
):
    """
    Send Outlook email with multiple inline snapshots and optional attachments.
    """

    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = ";".join(to_emails) if isinstance(to_emails, list) else to_emails
        if cc_emails:
            mail.CC = ";".join(cc_emails) if isinstance(cc_emails, list) else cc_emails

        mail.Subject = subject
        html_images = ""
        cid_counter = 1

        # --- Embed inline images ---
        for sheet_name, image_path in snapshots.items():
            attachment = mail.Attachments.Add(image_path)
            cid = f"MyId{cid_counter}"
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid)
            html_images += f"<img src='cid:{cid}'>"
            cid_counter += 1

        # --- download link ---
        download_html = (
            f'<p>Download link: <a href="{download_link}">{download_link}</a></p>'
            if download_link
            else ""
        )

        # --- file attachments ---
        if attachments:
            if not isinstance(attachments, list):
                attachments = [attachments]
            for file_path in attachments:
                if os.path.exists(file_path):
                    mail.Attachments.Add(file_path)
                    print(f"📎 Attached file: {file_path}")
                else:
                    print(f"⚠️ Attachment not found: {file_path}")

        # --- Compose HTML body ---
        mail.HTMLBody = f"""
        <html>
        <body>
            <p>Dear All,</p>
            <p>Greetings of the day!</p>
            <p>Please find the below <b>{report_name}</b>:</p>
            {html_images}
            {download_html}
            <p>Note: This is a system-generated report! For any assistance, kindly contact:</p>
            <p><a href="mailto:central.wfm-asr@kochartech.com">central.wfm-asr@kochartech.com</a></p>
            <p>Regards,</p>
            <p><b>WFM REPORTS (KocharTech)</b></p>
        </body>
        </html>
        """

        # --- Send Email ---
        mail.Send()
        print("✅ Email with inline snapshots and attachments sent successfully!")

    except Exception as e:
        print(f"❌ Failed to send email: {e}")


# ----------------- Example Usage -----------------

if __name__ == "__main__":
    try:
        file_path = (r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Sep'25\Raw Dump\Automation Output\ivr_hungupp.xlsb")
        output_dir = (r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Sep'25\Raw Dump\Automation Output")
        sheet_ranges = {
            "Sheet1": "A5:D13"
        }

        snapshots = take_excel_snapshots(file_path, sheet_ranges, output_dir)

        if snapshots: 
            send_snapshots_inline(
                to_emails = ["shankar.deshmukh2@emp.youbroadband.co.in"],                
                cc_emails = ["junaid.shamsi@maxicus.com", "chandrabhan.tomar@in.maxicus.com", "gurpreet.singh2@kochartech.com",
                             "archana@maxicus.com", "Central.MIS.Youbroadband@kochartech.com", "omprakash.sharma@in.maxicus.com", "amritpal.singh@maxicus.com"],                
                subject = "YouBroadband Hungup Report",
                snapshots=snapshots,
                report_name="YouBroadband IVR Hung-up Report",
                download_link="",
                attachments = r"\\172.17.52.16\172.17.3.195-data\KocharWFM\Zepto\Internal Dashboard\2025\Sep'25\Raw Dump\Automation Output\ivr_hungupp.xlsb"
            )

    except Exception as e:
        print(f"❌ Script failed: {e}")
        traceback.print_exc() 
        
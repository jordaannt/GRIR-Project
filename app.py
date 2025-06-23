import streamlit as st
import pandas as pd
import os
import tempfile
import shutil
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import GRIR  # Import the refactored GRIR analysis module
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from pathlib import Path

def create_temp_data_directory(temp_dir):
    """Create a temporary directory for data files within a given base temp dir."""
    data_dir = os.path.join(temp_dir, "data")
    os.makedirs(data_dir, exist_ok=True)
    return data_dir

def save_uploaded_file(uploaded_file, file_path):
    """Save uploaded file to the specified path."""
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

def send_custom_emails(summary_df, contacts_df, temp_dir, smtp_settings):
    """Send emails with custom SMTP settings."""
    issues_only = summary_df[summary_df['Action'].notna() & (summary_df['Action'] != "")]
    plant_po_issues = {}
    
    for plant, plant_group in issues_only.groupby('Plant'):
        report = f"<h2>GRIR Report {datetime.now().strftime('%B')} - {plant}</h2>"
        for po, po_group in plant_group.groupby('PO'):
            report += f"<br><b style='background-color: #FFFF00;'>PO {po}</b><br>"
            if po_group['Action'].str.startswith("Invoice has not been paid").all():
                report += f"<span style='color: red;'>{po_group['Action'].iloc[0]}</span><br><br>"
            else:
                for _, row in po_group.iterrows():
                    report += (
                        f"Line {int(row['Line'])} | {row['Line/Shade']} | {row['Description']}<br>"
                        f"You goods receipted: {int(row['Goods Receipt Qty'])} @ ${row['Goods Receipt Value']:.2f}<br>"
                        f"You were invoiced: {int(row['Invoice Qty'])} @ ${row['Invoice Receipt Value']:.2f}<br>"
                        f"<span style='color: red;'>{row['Action']}</span></b><br><br>"
                    )
        plant_po_issues[plant] = report

    # Send emails with custom settings
    for _, row in contacts_df.iterrows():
        plant, to_email = row['Plant'], row['Email']
        cc_email = row['CC'] if pd.notna(row['CC']) else ""
        message_body = plant_po_issues.get(plant)

        if message_body:
            plant_df = summary_df[summary_df['Plant'] == plant]
            plant_file = os.path.join(temp_dir, f"GRIR_Report_{plant}.xlsx")
            plant_df.to_excel(plant_file, index=False)
            GRIR.format_excel_file(plant_file)

            html_body = f"<html><body>{message_body}<p><i>Attached is the full GRIR report for your plant.</i></p></body></html>"
            msg = MIMEMultipart()
            msg['From'] = smtp_settings['sender_email']
            msg['To'] = to_email
            msg['CC'] = cc_email
            msg['Subject'] = f"GRIR Report – {plant}"
            msg.attach(MIMEText(html_body, 'html'))

            with open(plant_file, "rb") as f:
                part = MIMEApplication(f.read(), Name=f"GRIR_Report_{plant}.xlsx")
                part['Content-Disposition'] = f'attachment; filename="GRIR_Report_{plant}.xlsx"'
                msg.attach(part)
            
            recipients = [to_email] + ([cc_email] if cc_email else [])
            try:
                with smtplib.SMTP(smtp_settings['server'], smtp_settings['port']) as server:
                    server.starttls()
                    server.login(smtp_settings['sender_email'], smtp_settings['password'])
                    server.sendmail(smtp_settings['sender_email'], recipients, msg.as_string())
                st.success(f"✅ Sent GRIR summary to {to_email} (cc: {cc_email})")
            except Exception as e:
                st.error(f"❌ Failed to send email to {to_email}: {e}")
        else:
            st.info(f"No issues for plant {plant}, no email sent.")

def run_grir_analysis_in_temp_dir(ekbe_file, ekpo_file, email_file, price_tolerance, send_emails=False, smtp_settings=None):
    """Run the GRIR analysis using a temporary directory for file operations."""
    temp_dir = tempfile.mkdtemp()
    try:
        data_dir = create_temp_data_directory(temp_dir)
        
        # Define file paths
        ekbe_path = os.path.join(data_dir, "EKBE.xlsx")
        ekpo_path = os.path.join(data_dir, "EKPO.XLSX")
        email_path = os.path.join(data_dir, "email.xlsx")
        output_summary_path = os.path.join(temp_dir, "GRIR_summary.xlsx")
        
        # Save uploaded files to the temporary directory
        save_uploaded_file(ekbe_file, ekbe_path)
        save_uploaded_file(ekpo_file, ekpo_path)
        save_uploaded_file(email_file, email_path)
        
        # Run the analysis using the imported functions
        summary_df = GRIR.run_analysis(
            ekbe_path,
            ekpo_path,
            email_path,
            output_summary_path=output_summary_path,
            price_tolerance=price_tolerance,
            send_emails=False  # We'll handle emails separately
        )
        
        # Send custom emails if requested
        if send_emails and smtp_settings:
            contacts_df = pd.read_excel(email_path)
            send_custom_emails(summary_df, contacts_df, temp_dir, smtp_settings)
        
        # Collect generated report files
        plant_files = {}
        for file in os.listdir(temp_dir):
            if file.startswith("GRIR_Report_") and file.endswith(".xlsx"):
                plant_name = file.replace("GRIR_Report_", "").replace(".xlsx", "")
                plant_files[plant_name] = os.path.join(temp_dir, file)
        
        return summary_df, plant_files, temp_dir
    except Exception as e:
        # Ensure cleanup happens on error
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise e

def create_dashboard(summary_df):
    """Create interactive dashboard with charts and metrics."""
    st.subheader("Dashboard")
    
    # Key metrics
    total_pos = len(summary_df['PO'].unique())
    total_lines = len(summary_df)
    issues_df = summary_df[summary_df['Action'].notna() & (summary_df['Action'] != "")]
    issues_count = len(issues_df)
    
    # Create metrics row
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total POs", total_pos)
    col2.metric("Total Lines", total_lines)
    col3.metric("Issues Found", issues_count)
    issue_rate = (issues_count / total_lines) * 100 if total_lines > 0 else 0
    col4.metric("Issue Rate", f"{issue_rate:.1f}%")
    
    if issues_count > 0:
        st.subheader("Issue Breakdown")
        
        # Categorize issues for charts
        def categorize_issue(action):
            if "Invoice has not been paid" in action: return "Invoice Not Paid"
            if "higher quantity" in action or "lower quantity" in action: return "Quantity Mismatch"
            if "discrepancy" in action: return "Price Discrepancy"
            if "Short supply" in action: return "Short Supply/Credit"
            return "Other"

        # Fix the SettingWithCopyWarning by using .copy()
        issues_df = issues_df.copy()
        issues_df['Category'] = issues_df['Action'].apply(categorize_issue)
        issue_counts = issues_df['Category'].value_counts()
        
        c1, c2 = st.columns(2)
        with c1:
            fig_pie = px.pie(
                values=issue_counts.values,
                names=issue_counts.index,
                title="Issue Types Distribution"
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with c2:
            plant_issues = issues_df.groupby('Plant').size().sort_values(ascending=False)
            if not plant_issues.empty:
                fig_bar = px.bar(
                    plant_issues,
                    x=plant_issues.index,
                    y=plant_issues.values,
                    title="Issues by Plant",
                    labels={'x': 'Plant', 'y': 'Number of Issues'}
                )
                st.plotly_chart(fig_bar, use_container_width=True)

def generate_email_content_preview(summary_df, contacts_df):
    """Generate email content preview for each plant."""
    issues_only = summary_df[summary_df['Action'].notna() & (summary_df['Action'] != "")]
    plant_po_issues = {}
    
    for plant, plant_group in issues_only.groupby('Plant'):
        report = f"<h2>GRIR Report {datetime.now().strftime('%B')} - {plant}</h2>"
        for po, po_group in plant_group.groupby('PO'):
            report += f"<br><b style='background-color: #FFFF00;'>PO {po}</b><br>"
            if po_group['Action'].str.startswith("Invoice has not been paid").all():
                report += f"<span style='color: red;'>{po_group['Action'].iloc[0]}</span><br><br>"
            else:
                for _, row in po_group.iterrows():
                    report += (
                        f"Line {int(row['Line'])} | {row['Line/Shade']} | {row['Description']}<br>"
                        f"You goods receipted: {int(row['Goods Receipt Qty'])} @ ${row['Goods Receipt Value']:.2f}<br>"
                        f"You were invoiced: {int(row['Invoice Qty'])} @ ${row['Invoice Receipt Value']:.2f}<br>"
                        f"<span style='color: red;'>{row['Action']}</span></b><br><br>"
                    )
        plant_po_issues[plant] = report
    
    return plant_po_issues

def main():
    st.set_page_config(page_title="GRIR Analysis Dashboard", layout="wide")
    st.title("Run GR/IR Report")
    st.markdown("1. Run GR/IR report in SAP\n2. Copy all PO numbers\n3. Paste into SE16N 'EKPO' table."\
                 "Run and export to Excel.\n4. Paste into 'EKPE' table, run and export.\n\nDo not change any exported data.\n\nemail.xlsx should contain headers 'Plant', 'Email', 'CC'")
    
    # Sidebar for file uploads and settings
    with st.sidebar:
        st.header("Upload Files")
        ekbe_file = st.file_uploader("Upload EKBE data", type=['xlsx'])
        ekpo_file = st.file_uploader("Upload EKPO", type=['xlsx'])
        email_file = st.file_uploader("Upload email.xlsx", type=['xlsx'])
        
        st.markdown("---")
        st.header("Settings")
        price_tolerance = st.slider(
            "Price Tolerance ($)", 0.01, 1.00, 0.05, 0.01,
            help="Allowed difference between GR and IR values before flagging a price discrepancy."
        )
        
        st.markdown("---")
        st.header("Email Configuration")
        
        # Email settings
        enable_emails = st.checkbox("Enable Email Notifications", value=False)
        
        smtp_settings = None
        if enable_emails:
            email_mode = st.radio(
                "Email Configuration Mode",
                ["Use Default Settings", "Custom SMTP Settings"],
                help="Choose whether to use the default settings from GRIR.py or configure custom SMTP settings"
            )
            
            if email_mode == "Custom SMTP Settings":
                st.subheader("SMTP Configuration")
                smtp_settings = {
                    'server': st.text_input("SMTP Server", value="smtp.gmail.com"),
                    'port': st.number_input("SMTP Port", value=587, min_value=1, max_value=65535),
                    'sender_email': st.text_input("Sender Email", value=""),
                    'password': st.text_input("App Password", type="password", help="Use app password for Gmail")
                }
                
                if not all(smtp_settings.values()):
                    st.warning("⚠️ Please fill in all SMTP settings")
                    enable_emails = False
            else:
                # Use environment variables for default settings
                smtp_settings = {
                    'server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
                    'port': int(os.getenv('SMTP_PORT', '587')),
                    'sender_email': os.getenv('SENDER_EMAIL', 'duluxtradeops@gmail.com'),
                    'password': os.getenv('SENDER_PASSWORD', 'gljy uctw cfcn cqrz')
                }
                
                # Check if environment variables are set
                if smtp_settings['password'] == 'gljy uctw cfcn cqrz':
                    st.warning("Using default account.")
                else:
                    st.info("Using environment variables for email settings")
            
            # Email preview
            if email_file is not None:
                try:
                    contacts_df = pd.read_excel(email_file)
                    st.write("**Email Recipients:**")
                    for _, row in contacts_df.iterrows():
                        plant = row.get('Plant', 'Unknown')
                        email = row.get('Email', 'No email')
                        cc = row.get('CC', '')
                        cc_text = f" (CC: {cc})" if cc else ""
                        st.write(f"• {plant}: {email}{cc_text}")
                except Exception as e:
                    st.error(f"Error reading email file: {e}")
        
        run_analysis = st.button("Run", type="primary", disabled=not all([ekbe_file, ekpo_file, email_file]))
    
    # Main content area
    if run_analysis:
        with st.spinner("Running analysis... This may take a moment."):
            try:
                summary_df, plant_files, temp_dir = run_grir_analysis_in_temp_dir(
                    ekbe_file, ekpo_file, email_file, price_tolerance, 
                    send_emails=enable_emails, smtp_settings=smtp_settings
                )
                st.success("Report generation successful!")
                
                # Show email status if enabled
                if enable_emails:
                    st.success("Email send success!")
                
                create_dashboard(summary_df)
                
                st.subheader("Analysis Results")
                # Add filtering options here
                st.dataframe(summary_df, use_container_width=True, hide_index=True)
                
                # Display email content preview
                if email_file is not None:
                    try:
                        contacts_df = pd.read_excel(email_file)
                        plant_po_issues = generate_email_content_preview(summary_df, contacts_df)
                        
                        st.subheader("Email Content Preview")
                        st.markdown("Below is the content that would be sent to each plant:")
                        
                        for _, row in contacts_df.iterrows():
                            plant = row.get('Plant', 'Unknown')
                            email = row.get('Email', 'No email')
                            cc = row.get('CC', '')
                            cc_text = f" (CC: {cc})" if cc else ""
                            
                            with st.expander(f"{plant} - {email}{cc_text}"):
                                if plant in plant_po_issues:
                                    html_content = f"""
                                    <div style="font-family: Arial, sans-serif; line-height: 1.6;">
                                        {plant_po_issues[plant]}
                                    </div>
                                    """
                                    st.components.v1.html(html_content, height=400, scrolling=True)
                                else:
                                    st.info(f"No issues found for {plant} - no email would be sent.")
                    except Exception as e:
                        st.error(f"Error generating email preview: {e}")
                
                st.subheader("Download Reports")
                c1, c2 = st.columns(2)
                summary_excel_path = os.path.join(temp_dir, "GRIR_summary.xlsx")
                if os.path.exists(summary_excel_path):
                    with open(summary_excel_path, "rb") as f:
                        c1.download_button("Download Formatted Summary (Excel)", f.read(), "GRIR_Summary_Formatted.xlsx")
                
                for plant_name, file_path in plant_files.items():
                    with open(file_path, "rb") as f:
                        st.download_button(f"Download {plant_name} Report", f.read(), f"GRIR_Report_{plant_name}.xlsx")
                
                # Clean up temporary directory after all content is displayed
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception as e:
                st.error(f"❌ An error occurred during analysis: {e}")
                st.exception(e)
    else:
        st.info("Upload all required files and click 'Run' to begin.\n\nExpand the side bar on the left if it isn't visible")

if __name__ == "__main__":
    main() 
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, abort, send_file
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
import plotly.graph_objs as go
import re
import os

# List of YouTube video URLs
videos = [
    "https://www.youtube.com/embed/GMpBtwKKQI0?enablejsapi=1",
    "https://www.youtube.com/embed/mqh1ujofwsc?enablejsapi=1",
    "https://www.youtube.com/embed/xnRIDHTQC4Y?enablejsapi=1"
]

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'bramharoopasaraswati108')

# Get the sheet URL from environment variable
sheet_url = os.getenv('SHEET_URL', 'https://docs.google.com/spreadsheets/d/1ZrugwluaxqYKeLuuaImRtAyvqExEOB4eUEEf80titvo/edit?usp=sharing')

def send_email(subject, body):
    from_email = os.getenv('FROM_EMAIL', 'mitesh.k.blr@gmail.com')
    from_password = os.getenv('FROM_PASSWORD', 'jjxh dvgu benc cmor')

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = from_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, from_email, msg.as_string())
        server.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Function to access Google Sheets
def get_google_sheet(sheet_url):
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file('credentials.json', scopes=scopes)
    client = gspread.authorize(credentials)
    sheet = client.open_by_url(sheet_url).sheet1
    return sheet

sheet = get_google_sheet(sheet_url)
if sheet:
    print("Successfully connected to Google Sheets")
else:
    print("Connection to Google Sheets failed")


# Function to check if email already exists in Google Sheets
def email_exists(sheet, email):
    email_column = sheet.col_values(2)  # Assuming email is in the second column
    return email in email_column

# Function to store form data in Google Sheets
def store_in_google_sheet(sheet, data):
    sheet.append_row(data)

# Function to generate charts using Plotly
def generate_charts():
    placement_data = [94]
    avg_other_institutions_data = [61]
    successful_placements_data = [100, 120, 150, 180, 200, 160, 200, 200, 160, 80, 60, 200, 314, 420]

    combined_chart = go.Bar(
        x=['Placement Record', 'Other Institutions'],
        y=[placement_data[0], avg_other_institutions_data[0]],
        marker_color=['rgb(255, 99, 132)', 'rgb(54, 162, 235)']
    )
    combined_layout = go.Layout(
        title='Placement Record vs Other Institutions (%)',
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        titlefont=dict(color='#FFFFFF'),
        xaxis=dict(tickfont=dict(color='#FFFFFF')),
        yaxis=dict(tickfont=dict(color='#FFFFFF'))
    )
    combined_fig = go.Figure(data=[combined_chart], layout=combined_layout)
    combined_chart_div = combined_fig.to_html(full_html=False, default_height=400, default_width=400)

    successful_placements_chart = go.Scatter(
        x=['2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024'],
        y=successful_placements_data,
        mode='lines+markers',
        line=dict(color='rgb(255, 99, 71)', width=2),
        marker=dict(color='rgb(255, 99, 71)', size=8)
    )
    successful_placements_layout = go.Layout(
        title='Successful Subsequent Placements Over the Years',
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        titlefont=dict(color='#FFFFFF'),
        xaxis=dict(title='Year', tickfont=dict(color='#FFFFFF')),
        yaxis=dict(title='Number of Placements', tickfont=dict(color='#FFFFFF'))
    )
    successful_placements_fig = go.Figure(data=[successful_placements_chart], layout=successful_placements_layout)
    successful_placements_chart_div = successful_placements_fig.to_html(full_html=False, default_height=400, default_width=800)

    return {
        "placement_chart_div": combined_chart_div,
        "successful_placements_chart_div": successful_placements_chart_div,
    }

# Route to restrict download for logged-in users
@app.route('/protected_download/<path:filename>')
def protected_download(filename):
    if 'email' not in session:
        return redirect(url_for('index', error="Please register to download files"))
    
    allowed_files = ['fullstack.pdf', 'dsai.pdf', 'speakeasy.pdf', 'bfsi.pdf', 'supplychain.pdf', 'hrint.pdf', 'markint.pdf', 'schainint.pdf']
    
    if filename not in allowed_files:
        abort(404)
    
    try:
        return send_file(f'static/files/{filename}', as_attachment=True)
    except FileNotFoundError:
        abort(404)

# Home route to render the main page
@app.route('/')
def index():
    charts_div = generate_charts()
    error = request.args.get('error')
    return render_template('index.html', charts_div=charts_div, error=error, videos=videos)

@app.route('/submit_form', methods=['POST'])
def submit_form():
    if request.method == 'POST':
        try:
            # Retrieve form data
            fullname = request.form['fullName']
            email = request.form['email']
            phone_number = request.form['phoneNumber']
            program = request.form['program']

            # Validate email format
            if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
                return render_template('index.html', error="Invalid email format", charts_div=generate_charts(), videos=videos)

            # Access the Google Sheet
            sheet = get_google_sheet(sheet_url)
            
            # Check if the email already exists
            if email_exists(sheet, email):
                return render_template('index.html', error="Email already registered", charts_div=generate_charts(), videos=videos)

            # Prepare data to be stored in Google Sheets
            data = [fullname, email, phone_number, program]

            # Store data in Google Sheets
            store_in_google_sheet(sheet, data)

            # Save data in session
            session['fullname'] = fullname
            session['email'] = email
            session['phone_number'] = phone_number
            session['program'] = program

            # Redirect to thank you page
            return redirect(url_for('thank_you'))

        except Exception as e:
            print(f"Error: {e}")
            return render_template('index.html', error=str(e), charts_div=generate_charts(), videos=videos)

# Thank you page
@app.route('/thank_you')
def thank_you():
    # Retrieve data from session
    fullname = session.get('fullname')
    email = session.get('email')
    phone_number = session.get('phone_number')
    program = session.get('program')

    # Pass data to the thank_you template
    return render_template('thank_you.html', fullname=fullname, email=email, phone_number=phone_number, program=program)


# Route for logout
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)

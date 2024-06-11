from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import gspread
from google.oauth2.service_account import Credentials
import plotly.graph_objs as go
import re
import os

app = Flask(__name__)
app.secret_key = 'bramharoopasaraswati108'

# Function to send an email
def send_email(subject, body):
    from_email = 'mitesh.k.blr@gmail.com'
    from_password = 'jjxh dvgu benc cmor'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = 'mitesh.k.blr@gmail.com'
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, from_email, text)
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

# Function to check if email already exists in Google Sheets
def email_exists(sheet, email):
    email_column = sheet.col_values(2)  # Assuming the email is in the second column
    return email in email_column

# Function to store form data in Google Sheets
def store_in_google_sheet(sheet, data):
    sheet.append_row(data)

# Function to generate charts (code remains unchanged)
def generate_charts():
    # Data for charts (code remains unchanged)
    placement_data = [94]
    completion_rate_data = [95]
    fastest_placement_data = [10]
    mentors_data = [2000]
    techlabs_data = [3]
    placement_label = 'Placement Record'
    avg_other_institutions_data = [61]
    avg_other_institutions_label = 'Other Institutions'
    average_fastest_placement_data = [22]
    average_mentors_data = [1069]
    successful_placements_data = [100, 120, 150, 180, 200, 160, 200, 200, 160, 80, 60, 200, 314, 420]

    # Create the placement chart
    combined_chart = go.Bar(x=[placement_label, avg_other_institutions_label], y=[placement_data[0], avg_other_institutions_data[0]], marker_color=['rgb(255, 99, 132)', 'rgb(54, 162, 235)'])
    combined_layout = go.Layout(title='Placement Record vs Other Institutions (%)', titlefont=dict(color='#FFFFFF'), plot_bgcolor='#1D1336', paper_bgcolor='#1D1336', xaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')), yaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')))
    combined_fig = go.Figure(data=[combined_chart], layout=combined_layout)
    combined_chart_div = combined_fig.to_html(full_html=False, default_height=400, default_width=400)

    completion_rate_chart = go.Pie(
        labels=['Average Training Completion Rate (%)'],
        values=completion_rate_data,
        hole=0.5
    )

    completion_rate_layout = go.Layout(
        title='Average Training Completion Rate (%)',
        titlefont=dict(color='#FFFFFF'),
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        xaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')),
        yaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF'))
    )

    completion_rate_fig = go.Figure(data=[completion_rate_chart], layout=completion_rate_layout)
    completion_rate_fig.update_layout(
        height=400,
        width=400,
        showlegend=False,
    )

    completion_rate_chart_div = completion_rate_fig.to_html(full_html=False)

    # Create the fastest placement chart
    fastest_placement_chart = go.Bar(
        x=['Expertrons', 'Avg of others'],
        y=fastest_placement_data + average_fastest_placement_data,
        marker_color=['rgb(75, 192, 192)', 'rgb(255, 99, 71)']
    )

    fastest_placement_layout = go.Layout(
        title='Fastest Placement Time (Days)',
        titlefont=dict(color='#FFFFFF'),
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        xaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')),
        yaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF'))
    )

    fastest_placement_fig = go.Figure(data=[fastest_placement_chart], layout=fastest_placement_layout)
    fastest_placement_chart_div = fastest_placement_fig.to_html(full_html=False, default_height=400, default_width=400)

    # Create the mentors chart with two bars
    mentors_chart = go.Bar(
        x=['Expertrons mentors', 'Average Mentors'],
        y=mentors_data + average_mentors_data,
        marker_color=['rgb(255, 206, 86)', 'rgb(255, 99, 71)']
    )

    mentors_layout = go.Layout(
        title='Number of Mentors',
        titlefont=dict(color='#FFFFFF'),
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        xaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')),
        yaxis=dict(titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF'))
    )

    mentors_fig = go.Figure(data=[mentors_chart], layout=mentors_layout)
    mentors_chart_div = mentors_fig.to_html(full_html=False, default_height=400, default_width=400)

    # Create the line chart for successful subsequent placements
    successful_placements_chart = go.Scatter(
        x=['2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023', '2024'],
        y=successful_placements_data,
        mode='lines+markers',
        line=dict(color='rgb(255, 99, 71)', width=2),
        marker=dict(color='rgb(255, 99, 71)', size=8),
        name='Successful Subsequent Placements'
    )

    successful_placements_layout = go.Layout(
        title='Successful Subsequent Placements Over the Years',
        titlefont=dict(color='#FFFFFF'),
        plot_bgcolor='#1D1336',
        paper_bgcolor='#1D1336',
        xaxis=dict(title='Year', titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF')),
        yaxis=dict(title='Number of Placements', titlefont=dict(color='#FFFFFF'), tickfont=dict(color='#FFFFFF'))
    )

    successful_placements_fig = go.Figure(data=[successful_placements_chart], layout=successful_placements_layout)
    successful_placements_chart_div = successful_placements_fig.to_html(full_html=False, default_height=400, default_width=800)

    return {
        "placement_chart_div": combined_chart_div,
        "completion_rate_chart_div": completion_rate_chart_div,
        "fastest_placement_chart_div": fastest_placement_chart_div,
        "mentors_chart_div": mentors_chart_div,
        "successful_placements_chart_div": successful_placements_chart_div,
    }

# Define routes
@app.route('/')
def index():
    charts_div = generate_charts()
    return render_template('index.html', charts_div=charts_div)

@app.route('/submit_form', methods=['POST'])
def submit_form():
    if request.method == 'POST':
        # Retrieve form data
        fullname = request.form['fullName']
        email = request.form['email']
        preferred_domain = request.form['preferred_domain']
        career_stage = request.form['career_stage']
        birthdate = request.form['birthdate']
        preferred_time = request.form['preferred_time']
        
        # Validate email format
        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            charts_div = generate_charts()
            return render_template('index.html', error="Invalid email format", charts_div=charts_div)
        
        # Access Google Sheets
        sheet_url = 'https://docs.google.com/spreadsheets/d/1ZrugwluaxqYKeLuuaImRtAyvqExEOB4eUEEf80titvo/edit?usp=sharing'
        sheet = get_google_sheet(sheet_url)
        
        # Check if email already exists
        if email_exists(sheet, email):
            charts_div = generate_charts()
            return render_template('index.html', error="Email already registered", charts_div=charts_div)

        # Store data in session
        session['fullname'] = fullname
        session['email'] = email
        session['preferred_domain'] = preferred_domain
        session['career_stage'] = career_stage
        session['birthdate'] = birthdate
        session['preferred_time'] = preferred_time
        
        # Store form data in Google Sheets
        data = [fullname, email, preferred_domain, career_stage, birthdate, preferred_time]
        store_in_google_sheet(sheet, data)
        
        # Send confirmation email
        subject = "New Form Submission"
        body = f"Form Submission Details:\n\nFull Name: {fullname}\nEmail: {email}\nPreferred Domain: {preferred_domain}\nCareer Stage: {career_stage}\nBirthdate: {birthdate}\nPreferred Time: {preferred_time}"
        send_email(subject, body)
        
        # Redirect to the "Thank You" page with extracted details
        return redirect(url_for('thank_you'))

@app.route('/thank_you')
def thank_you():
    data = {
        'fullname': session.get('fullname'),
        'email': session.get('email'),
        'preferred_domain': session.get('preferred_domain'),
        'career_stage': session.get('career_stage'),
        'birthdate': session.get('birthdate'),
        'preferred_time': session.get('preferred_time')
    }
    return render_template('thank_you.html', **data)

@app.route('/download_file/<path:filename>')
def download_file(filename):
    return send_from_directory(directory='files', filename=filename)

@app.route('/end')
def end():
    return render_template('end.html')

@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True)

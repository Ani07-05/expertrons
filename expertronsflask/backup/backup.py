from flask import Flask, render_template, request, redirect, url_for
import sqlite3
import plotly.graph_objs as go

app = Flask(__name__)

# Function to create a SQLite database connection
def create_connection():
    conn = None
    try:
        conn = sqlite3.connect('form_data.db')
        return conn
    except sqlite3.Error as e:
        print(e)
    return conn

# Function to create a table in the database
def create_table(conn):
    try:
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS form_data
                          (id INTEGER PRIMARY KEY AUTOINCREMENT,
                           fullname TEXT,
                           email TEXT,
                           preferred_domain TEXT,
                           career_stage TEXT,
                           birthdate TEXT,
                           time TEXT)''')
        conn.commit()
    except sqlite3.Error as e:
        print(e)

def create_registrations_table(conn):
    try:
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS registrations_data
                          (id INTEGER PRIMARY KEY AUTOINCREMENT,
                           fullname TEXT,
                           email TEXT,
                           program TEXT,
                           sdomain TEXT,
                           otherContact TEXT)''')
    except sqlite3.Error as e:
        print(e)

# Function to generate charts
def generate_charts():
    # Data for charts
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
    successful_placements_data = [100, 120, 150, 180, 200, 160, 200, 200, 160, 80, 60, 200, 314,420]

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
        y=fastest_placement_data + average_fastest_placement_data,  # Concatenate both datasets
        marker_color=['rgb(75, 192, 192)', 'rgb(255, 99, 71)']  # Assign different colors to bars
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
        y=mentors_data + average_mentors_data,  # Concatenate both datasets
        marker_color=['rgb(255, 206, 86)', 'rgb(255, 99, 71)']  # Assign different colors to bars
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
        x=['2010', '2011', '2012', '2013', '2014', '2015', '2016', '2017', '2018', '2019', '2020', '2021','2022','2023','2024'],
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
        time = request.form['time']
        
        # Save form data to the database
        conn = create_connection()
        if conn is not None:
            create_registrations_table(conn)
            cursor = conn.cursor()
            cursor.execute('''INSERT INTO form_data
                              (fullname, email, preferred_domain, career_stage, birthdate, time)
                              VALUES (?, ?, ?, ?, ?, ?)''', (fullname, email, preferred_domain, career_stage, birthdate, time))
            conn.commit()
            conn.close()
        
        # Redirect to the "Thank You" page with extracted details
        return redirect(url_for('thank_you', fullname=fullname, email=email,
                                preferred_domain=preferred_domain, career_stage=career_stage,
                                birthdate=birthdate, time=time))

@app.route('/submit_registration_form', methods=['POST'])
def submit_registration_form():
    if request.method == 'POST':
        # Retrieve form data
        fullname = request.form['fullname']
        email = request.form['email']
        program = request.form['program']
        sdomain=request.form['interests']
        otherContact=request.form['otherContact']
        
        # Save form data to the database
        conn = create_connection()
        if conn is not None:
            create_registrations_table(conn)  # Create the table if it doesn't exist
            cursor = conn.cursor()
            cursor.execute('''INSERT INTO registrations_data
                              (fullname, email, program, sdomain, otherContact)
                              VALUES (?, ?, ?, ?, ?)''', (fullname, email, program, sdomain, otherContact))
            conn.commit()
            conn.close()
        
    return redirect(url_for('registration', fullname=fullname, email=email,
                            program=program, sdomain=sdomain, otherContact=otherContact))


@app.route('/thank_you')
def thank_you():
    # Extract details from URL parameters
    fullname = request.args.get('fullname')
    email = request.args.get('email')
    preferred_domain = request.args.get('preferred_domain')
    career_stage = request.args.get('career_stage')
    birthdate = request.args.get('birthdate')
    time = request.args.get('time')
    
    return render_template('thank_you.html', fullname=fullname, email=email,
                           preferred_domain=preferred_domain, career_stage=career_stage,
                           birthdate=birthdate, time=time)
@app.route('/registration')
def registration():
    # Extract details from URL parameters
    fullname = request.args.get('fullname')
    email = request.args.get('email')
    program = request.args.get('program')
    sdomain = request.args.get('sdomain')
    otherContact = request.args.get('otherContact')
    
    return render_template('registration.html', fullname=fullname, email=email,
                           program=program, sdomain=sdomain, otherContact=otherContact)

# Error handling for 404 page not found
@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html'), 404

if __name__ == '__main__':
    app.run(debug=True)

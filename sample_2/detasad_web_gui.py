from flask import Flask, send_file
from flask import render_template, request, redirect

#, render_template, request, redirect
import configparser
import subprocess

app = Flask(__name__)

# Load the initial configuration
config = configparser.ConfigParser()
config.read('DetaSadMonitoring.cfg')

@app.route('/')
def index():
    #return send_file('./index.html')
    return app.send_static_file('index.html')
   # return render_template('./index.html', config=config)

@app.route('/', methods=['POST'])
def update_config():
    if request.method == 'POST':
        # Update the configuration with form data
        config['form_to']['start_date'] = request.form['start_date']
        config['form_to']['end_date'] = request.form['end_date']
        config['email']['mail_to'] = request.form['mail_to']
        config['threshold']['incident_threshold'] = request.form['incident_threshold']
        config['threshold']['frequency_window'] = request.form['frequency_window']

        # Write the updated configuration to the .cfg file
        with open('DetaSadMonitoring.cfg', 'w') as configfile:
            config.write(configfile)

        # Execute the test script
        subprocess.run(['python', 'DetaSadMonitoring.py'])

        return redirect('/')

if __name__ == '__main__':
    app.run(debug=True)
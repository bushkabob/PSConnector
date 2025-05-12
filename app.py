import threading
from flask import Flask, request, jsonify
from flask_socketio import SocketIO, emit
import win32com.client.dynamic

from methods import ComMethods
from events import SocketIOEvents, EventsWithCOM
import win32com.client
import pythoncom

app = Flask(__name__)
socketio = SocketIO(app)

prog_id = "Commissure.PACSConnector.RadWhereCOM"
ps = win32com.client.dynamic.Dispatch(prog_id)

com_methods = ComMethods(ps)

@app.route('/login', methods=['POST'])
def login():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    server = data.get('server')
    return jsonify({"message": com_methods.Login(username, password, server)})

@app.route('/loginex', methods=['POST'])
def login_ex():
    data = request.get_json()
    username = data.get('username')
    password = data.get('password')
    return jsonify({"message": com_methods.LoginEx(username, password)})

@app.route('/logout', methods=['POST'])
def logout():
    return jsonify({"message": com_methods.Logout()})

@app.route('/logoutex', methods=['POST'])
def logout_ex():
    return jsonify({"message": com_methods.LogoutEx()})

@app.route('/start', methods=['POST'])
def start():
    data = request.get_json()
    started = data.get('started')
    return jsonify({"message": com_methods.Start(started)})

@app.route('/terminate', methods=['POST'])
def terminate():
    return jsonify({"message": com_methods.Terminate()})

@app.route('/stop', methods=['POST'])
def stop():
    return jsonify({"message": com_methods.Stop()})

@app.route('/createreport', methods=['POST'])
def create_report():
    data = request.get_json()
    accession = data.get('accession')
    bStartDictation = data.get('bStartDictation')
    return jsonify({"message": com_methods.CreateNewReport(accession, bStartDictation)})

@app.route('/openreport', methods=['POST'])
def open_report():
    data = request.get_json()
    accessions = data.get('accessions')
    site = data.get('site', "")
    return jsonify({"message": com_methods.OpenReport(accessions, site)})

@app.route('/closereport', methods=['POST'])
def close_report():
    data = request.get_json()
    shouldSign = data.get('shouldSign')
    markPrelim = data.get('markPrelim')
    return jsonify({"message": com_methods.CloseReport(shouldSign, markPrelim)})

@app.route('/savereport', methods=['POST'])
def save_report():
    data = request.get_json()
    shouldClose = data.get('shouldClose')
    return jsonify({"message": com_methods.SaveReport(shouldClose)})

@app.route('/insertautotext', methods=['POST'])
def insert_autotext():
    data = request.get_json()
    autoTextName = data.get('autoTextName')
    strToReplace = data.get('strToReplace')
    return jsonify({"message": com_methods.InsertAutoText(autoTextName, strToReplace)})

@app.route('/previeworder', methods=['POST'])
def preview_order():
    data = request.get_json()
    accessionNumbers = data.get('accessionNumbers')
    site = data.get('site', "")
    return jsonify({"message": com_methods.PreviewOrder(accessionNumbers, site)})

@app.route('/associateorders', methods=['POST'])
def associate_orders():
    data = request.get_json()
    accessionNumbers = data.get('accessionNumbers')
    return jsonify({"message": com_methods.AssociateOrders(accessionNumbers)})

@app.route('/associateorderscurrent', methods=['POST'])
def associate_orders_current():
    data = request.get_json()
    currentAccession = data.get('currentAccession')
    newAccessions = data.get('newAccessions')
    site = data.get('site', "")
    return jsonify({"message": com_methods.AssociateOrdersWCurrent(currentAccession, newAccessions, site)})

@app.route('/dissociateorders', methods=['POST'])
def dissociate_orders():
    data = request.get_json()
    accessionNumbers = data.get('accessionNumbers')
    return jsonify({"message": com_methods.DissociateOrders(accessionNumbers)})

@app.route('/getactiveaccessions', methods=['GET'])
def get_active_accessions():
    return jsonify({"accessions": com_methods.GetActiveAccessions()})

@app.route('/getalwaysontop', methods=['GET'])
def get_always_on_top():
    return jsonify({"alwaysOnTop": com_methods.GetAlwaysOnTop()})

@app.route('/getminimized', methods=['GET'])
def get_minimized():
    return jsonify({"minimized": com_methods.GetMinimized()})

@app.route('/getrestrictedsession', methods=['GET'])
def get_restricted_session():
    return jsonify({"restrictedSession": com_methods.GetRestrictedSession()})

@app.route('/getrestrictedworkflow', methods=['GET'])
def get_restricted_workflow():
    return jsonify({"restrictedWorkflow": com_methods.GetRestrictedWorkflow()})

@app.route('/getsitename', methods=['GET'])
def get_site_name():
    return jsonify({"siteName": com_methods.GetSiteName()})

@app.route('/getusername', methods=['GET'])
def get_username():
    return jsonify({"username": com_methods.GetUsername()})

@app.route('/getloggedin', methods=['GET'])
def get_logged_in():
    return jsonify({"loggedIn": com_methods.GetLoggedIn()})

@app.route('/getsite', methods=['GET'])
def get_site():
    return jsonify({"site": com_methods.GetSite()})

@app.route('/getvisible', methods=['GET'])
def get_visible():
    return jsonify({"visible": com_methods.GetVisible()})

@app.route('/setalwaysontop', methods=['POST'])
def set_always_on_top():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetAlwaysOnTop(value)
    return jsonify({"message": "AlwaysOnTop set successfully"})

@app.route('/setminimized', methods=['POST'])
def set_minimized():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetMinimized(value)
    return jsonify({"message": "Minimized set successfully"})

@app.route('/setrestrictedsession', methods=['POST'])
def set_restricted_session():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetRestrictedSession(value)
    return jsonify({"message": "RestrictedSession set successfully"})

@app.route('/setrestrictedworkflow', methods=['POST'])
def set_restricted_workflow():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetRestrictedWorkflow(value)
    return jsonify({"message": "RestrictedWorkflow set successfully"})

@app.route('/setsitename', methods=['POST'])
def set_site_name():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetSite(value)
    return jsonify({"message": "Site set successfully"})

@app.route('/setvisible', methods=['POST'])
def set_visible():
    data = request.get_json()
    value = data.get('value')
    com_methods.SetVisible(value)
    return jsonify({"message": "Visible set successfully"})

@socketio.on('connect')
def handle_connect():
    print("Client connected")
    emit('status', {'message': 'Connected to WebSocket'})

@socketio.on('disconnect')
def handle_disconnect():
    print("Client disconnected")

@socketio.on('send_report')
def handle_send_report(data):
    print(f"Received report data: {data}")
    emit('status', {'message': 'Report data received successfully'})

if __name__ == '__main__':
    def pump_com_events():
        socketio_events = SocketIOEvents(socketio)
        eventsWithCom: EventsWithCOM = win32com.client.WithEvents(ps, EventsWithCOM)
        eventsWithCom.configure(socketio_events)
        while True:
            pythoncom.PumpWaitingMessages()

    com_thread = threading.Thread(target=pump_com_events)
    com_thread.daemon = True
    com_thread.start()

    # Run both CRUD and WebSocket with Flask-SocketIO
    socketio.run(app, debug=True)
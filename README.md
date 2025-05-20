# PSConnector

**PSConnector** is a Flask-based REST and WebSocket API interface for interacting with the RadWhere COM interface provided by PowerScribeOne via Python on Windows. It enables communication between a REST HTTP server and the RadWhere reporting environment through COM, with real-time updates powered by Socket.IO.

## Features

- RESTful endpoints for logging in, managing reports, and accessing user/session/system info.
- WebSocket integration for real-time event broadcasting (e.g., report opened, user logged in).
- Background COM event listening using `win32com.client`.

## Prerequisites

- Windows OS w/ RadWhere COM component (provided by PowerScribe)
- Python 3.7+
- Dependencies:
  - Flask
  - Flask-SocketIO
  - pywin32

Install dependencies:

```bash
pip install flask flask-socketio pywin32
```

## Getting Started

1. Clone the repository.
2. Ensure the COM component `Commissure.PACSConnector.RadWhereCOM` is installed and registered.
3. Run the application:

```bash
python app.py
```

The server will start and listen on `http://127.0.0.1:5000/`.

## API Overview

### Authentication

- `POST /login`
- `POST /loginex`
- `POST /logout`
- `POST /logoutex`

### Report Handling

- `POST /createreport`
- `POST /openreport`
- `POST /closereport`
- `POST /savereport`

### Workflow & State

- `POST /start`
- `POST /terminate`
- `POST /stop`

### Order Management

- `POST /previeworder`
- `POST /associateorders`
- `POST /associateorderscurrent`
- `POST /dissociateorders`

### Properties (GET/SET)

- `GET /getactiveaccessions`
- `GET /getalwaysontop`
- `GET /getminimized`
- `GET /getrestrictedsession`
- `GET /getrestrictedworkflow`
- `GET /getsitename`
- `GET /getusername`
- `GET /getloggedin`
- `GET /getsite`
- `GET /getvisible`

- `POST /setalwaysontop`
- `POST /setminimized`
- `POST /setrestrictedsession`
- `POST /setrestrictedworkflow`
- `POST /setsitename`
- `POST /setvisible`

### WebSocket Events

- `report_event`: `opened`, `closed`, `changed`
- `user_event`: `logged_in`, `logged_out`
- `error_event`: `occurred`
- `prefetch_event`: `requested`
- `termination_event`, `dictation_event`, etc.

## Notes

- Designed for integration with dictation systems using RadWhere (ie. PowerScribe).
- The COM interface must be available and accessible from the machine running this application.
- Be aware that `win32com.client` may require administrative permissions.

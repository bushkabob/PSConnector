import win32com.client
import threading
import queue
from flask import Flask
from flask_socketio import SocketIO, emit


class SocketIOEvents:
    def __init__(self, socketio: SocketIO):
        self.socketio = socketio
        self.event_queue = queue.Queue()

        self.event_thread = threading.Thread(target=self._process_events)
        self.event_thread.daemon = True
        self.event_thread.start()

    def emit_report_event(self, event_name, **kwargs):
        """Emit events to SocketIO clients."""
        self.socketio.emit(event_name, kwargs)

    def _process_events(self):
        """Process events in the background and emit them to clients."""
        while True:
            event = self.event_queue.get()
            if event is None:
                break
            event_name, event_data = event
            self.emit_report_event(event_name, **event_data)

    def add_event_to_queue(self, event_name, **kwargs):
        """Add an event to the queue for processing by the background thread."""
        self.event_queue.put((event_name, kwargs))


class EventsWithCOM:
    socketio_events = None
    ps_events = None
    
    def configure(cls, socketio_events: SocketIOEvents):
        cls.socketio_events = socketio_events

    def OnReportOpened(self, site, accession, status=None, isAddendum=None, plainText=None, richText=None):
        print(f"Report opened: {site}, {accession}")
        self.socketio_events.add_event_to_queue('report_event', event='opened', site=site, accession=accession,
                                                status=status, is_addendum=isAddendum,
                                                plain_text=plainText, rich_text=richText)

    def OnReportClosed(self, site, accession, status=None, isAddendum=None, plainText=None, richText=None):
        print(f"Report closed: {site}, {accession}")
        self.socketio_events.add_event_to_queue('report_event', event='closed', site=site, accession=accession,
                                                status=status, is_addendum=isAddendum,
                                                plain_text=plainText, rich_text=richText)

    def OnReportChanged(self, site, accession, status=None, isAddendum=None, plainText=None, richText=None):
        print(f"Report changed: {site}, {accession}")
        self.socketio_events.add_event_to_queue('report_event', event='changed', site=site, accession=accession,
                                                status=status, is_addendum=isAddendum,
                                                plain_text=plainText, rich_text=richText)

    def OnUserLoggedIn(self, username):
        print(f"User logged in: {username}")
        self.socketio_events.add_event_to_queue('user_event', event='logged_in', username=username)

    def OnUserLoggedOut(self, username):
        print(f"User logged out: {username}")
        self.socketio_events.add_event_to_queue('user_event', event='logged_out', username=username)

    def OnErrorOccurred(self, errCode, message):
        print(f"Error occurred: {errCode}, {message}")
        self.socketio_events.add_event_to_queue('error_event', event='occurred', err_code=errCode, message=message)

    def OnPrefetchRequested(self, site, accessionNumbers):
        print(f"Prefetch requested: {site}, {accessionNumbers}")
        self.socketio_events.add_event_to_queue('prefetch_event', event='requested', site=site, accession_numbers=accessionNumbers)

    def OnTerminated(self):
        print("Program terminated")
        self.socketio_events.add_event_to_queue('termination_event', event='terminated')

    def OnDictationStarted(self):
        print("Dictation started")
        self.socketio_events.add_event_to_queue('dictation_event', event='started')

    def OnDictationStopped(self):
        print("Dictation stopped")
        self.socketio_events.add_event_to_queue('dictation_event', event='stopped')

    def OnAudioTranscribed(self, textRecognized):
        print(f"Audio transcribed: {textRecognized}")
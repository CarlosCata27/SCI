import pyrebase as fb
from Pyrebase.config import firebaseConfig

class FirebaseStorage:

    def __init__(self):
        self.firebase = fb.initialize_app(firebaseConfig)
        self.storage = self.firebase.storage()

    def upload_file(self, local_filename, cloud_filename):
        self.storage.child(cloud_filename).put(local_filename)

    def get_file_url(self, cloud_filename):
        return self.storage.child(cloud_filename).get_url(None)
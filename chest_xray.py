import warnings
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QPushButton, QFileDialog
from PyQt5.QtGui import QMovie
from keras.models import load_model
from keras.preprocessing import image
from keras.applications.vgg16 import preprocess_input
import numpy as np
from win32com.client import Dispatch

warnings.filterwarnings('ignore')

class PneumoniaApp(QMainWindow):
    def __init__(self):
        super(PneumoniaApp, self).__init__()

        self.setGeometry(0, 0, 701, 611)
        self.setWindowTitle("PNEUMONIA Detection Apps")

        self.central_widget = QWidget(self)
        self.central_widget.setStyleSheet("background-color: #035874;")

        self.frame = QWidget(self.central_widget)
        self.frame.setGeometry(0, 0, 701, 611)

        self.label = QLabel(self.frame)
        self.label.setGeometry(80, -60, 541, 561)
        self.label.setText("")
        self.gif = QMovie("picture.gif")
        self.label.setMovie(self.gif)
        self.gif.start()

        self.label_2 = QLabel(self.frame)
        self.label_2.setGeometry(140, 430, 591, 41)
        self.label_2.setText("Chest X_ray PNEUMONIA Detection")
        self.label_2.setStyleSheet("font-size: 24px; font-weight: bold; color: white;")

        self.pushButton = QPushButton(self.frame)
        self.pushButton.setGeometry(30, 530, 201, 31)
        self.pushButton.setText("Upload Image")
        self.pushButton.setStyleSheet("QPushButton{border-radius: 10px; background-color:#DF582C;}"
                                      "QPushButton:hover {background-color: #7D93E0;}")
        self.pushButton.clicked.connect(self.upload_image)

        self.pushButton_2 = QPushButton(self.frame)
        self.pushButton_2.setGeometry(450, 530, 201, 31)
        self.pushButton_2.setText("Prediction")
        self.pushButton_2.setStyleSheet("QPushButton{border-radius: 10px; background-color:#DF582C;}"
                                        "QPushButton:hover {background-color: #7D93E0;}")
        self.pushButton_2.clicked.connect(self.predict_result)

        self.result_label = QLabel(self.frame)
        self.result_label.setGeometry(200, 500, 300, 31)
        self.result_label.setStyleSheet("font-size: 30px; font-weight: bold; color:#DF582C;")

        self.setCentralWidget(self.central_widget)

        self.result = None

    def upload_image(self):
        filename, _ = QFileDialog.getOpenFileName()
        path = str(filename)
        print(path)

        model = load_model('chest_xray.h5')
        img_file = image.load_img(path, target_size=(224, 224))
        x = image.img_to_array(img_file)
        x = np.expand_dims(x, axis=0)
        img_data = preprocess_input(x)
        self.result = model.predict(img_data)

    def predict_result(self):
        if self.result is not None:
            print(self.result)
            if self.result[0][0] > 0.5:
                print("Result is Normal")
                self.speak("Result is Normal")
                self.pushButton.hide()
                self.pushButton_2.hide()
                self.result_label.setText("Result is Normal")
            else:
                print("Affected By PNEUMONIA")
                self.speak("Affected By PNEUMONIA")
                self.pushButton.hide()
                self.pushButton_2.hide()
                self.result_label.setText("Affected By PNEUMONIA")

          

            # Process events to update UI immediately
            QApplication.processEvents()

    def speak(self, text):
        speak = Dispatch(("SAPI.SpVoice"))
        speak.Speak(text)


if __name__ == "__main__":
    app = QApplication([])
    pneumonia_app = PneumoniaApp()
    pneumonia_app.show()
    app.exec_()

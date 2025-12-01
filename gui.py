import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QFileDialog, QMessageBox)
from PyQt6.QtCore import Qt
import comment  # Import the comment module

class CommenterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Document Commenter UI')
        self.setGeometry(300, 300, 600, 250)

        layout = QVBoxLayout()

        # --- Input Document Selection ---
        layout.addWidget(QLabel('Input Document (.docx or .pdf):'))
        
        doc_layout = QHBoxLayout()
        self.entry_doc = QLineEdit()
        doc_layout.addWidget(self.entry_doc)
        
        btn_browse_doc = QPushButton('Browse')
        btn_browse_doc.clicked.connect(self.browse_doc)
        doc_layout.addWidget(btn_browse_doc)
        
        layout.addLayout(doc_layout)

        # --- Annotation JSON Selection ---
        layout.addWidget(QLabel('Annotations JSON Source:'))
        
        json_layout = QHBoxLayout()
        self.entry_json = QLineEdit()
        
        # Set default to annotations-example.json if it exists
        default_json = os.path.abspath("annotations-example.json")
        if os.path.exists(default_json):
            self.entry_json.setText(default_json)
            
        json_layout.addWidget(self.entry_json)
        
        btn_browse_json = QPushButton('Browse')
        btn_browse_json.clicked.connect(self.browse_json)
        json_layout.addWidget(btn_browse_json)
        
        layout.addLayout(json_layout)

        # --- Run Button ---
        self.btn_run = QPushButton('Run Commenter')
        self.btn_run.clicked.connect(self.run_commenter)
        self.btn_run.setMinimumHeight(40)
        layout.addWidget(self.btn_run)

        # --- Status Label ---
        self.lbl_status = QLabel('Ready')
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_status)

        self.setLayout(layout)

    def browse_doc(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Input Document", 
            "", 
            "Documents (*.docx *.pdf);;Word Document (*.docx);;PDF Document (*.pdf);;All Files (*)"
        )
        if file_path:
            self.entry_doc.setText(file_path)

    def browse_json(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Annotations JSON", 
            "", 
            "JSON Files (*.json);;All Files (*)"
        )
        if file_path:
            self.entry_json.setText(file_path)

    def run_commenter(self):
        doc_path = self.entry_doc.text().strip()
        json_path = self.entry_json.text().strip()

        if not doc_path:
            QMessageBox.critical(self, "Error", "Please select an input document.")
            return
        if not os.path.exists(doc_path):
            QMessageBox.critical(self, "Error", f"Input document not found: {doc_path}")
            return
            
        if not json_path:
            QMessageBox.critical(self, "Error", "Please select an annotations JSON file.")
            return
        if not os.path.exists(json_path):
            QMessageBox.critical(self, "Error", f"JSON file not found: {json_path}")
            return

        try:
            self.lbl_status.setText("Running...")
            QApplication.processEvents()

            # Logic lifted from comment.py main()
            annotations = comment.load_annotations(json_path)
            
            base, ext = os.path.splitext(doc_path)
            out_path = f"{base}-annotated{ext}"
            
            ext_lower = ext.lower()
            if ext_lower == ".docx":
                comment.annotate_docx(doc_path, out_path, annotations)
            elif ext_lower == ".pdf":
                comment.annotate_pdf(doc_path, out_path, annotations)
            else:
                raise ValueError("Unsupported file extension. Only .docx and .pdf are supported.")
            
            self.lbl_status.setText(f"Success! Saved to: {os.path.basename(out_path)}")
            QMessageBox.information(self, "Success", f"Wrote annotated file to:\n{out_path}")
            
        except Exception as e:
            self.lbl_status.setText("Error occurred")
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = CommenterApp()
    ex.show()
    sys.exit(app.exec())

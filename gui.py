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
        self.setGeometry(300, 300, 600, 400)

        layout = QVBoxLayout()

        # --- OpenAI API Key ---
        layout.addWidget(QLabel('OpenAI API Key:'))
        self.entry_api_key = QLineEdit()
        self.entry_api_key.setEchoMode(QLineEdit.EchoMode.Password)
        # Try to load from env
        env_key = os.environ.get("OPENAI_API_KEY", "")
        if env_key:
            self.entry_api_key.setText(env_key)
        layout.addWidget(self.entry_api_key)

        # --- Input Document Selection ---
        layout.addWidget(QLabel('Input Document (Paper to Grade):'))
        
        doc_layout = QHBoxLayout()
        self.entry_doc = QLineEdit()
        doc_layout.addWidget(self.entry_doc)
        
        btn_browse_doc = QPushButton('Browse')
        btn_browse_doc.clicked.connect(lambda: self.browse_file(self.entry_doc, "Select Input Document", "Documents (*.docx *.pdf)"))
        doc_layout.addWidget(btn_browse_doc)
        
        layout.addLayout(doc_layout)

        # --- Rubric Selection ---
        layout.addWidget(QLabel('Rubric (Optional):'))
        rubric_layout = QHBoxLayout()
        self.entry_rubric = QLineEdit()
        rubric_layout.addWidget(self.entry_rubric)
        btn_browse_rubric = QPushButton('Browse')
        btn_browse_rubric.clicked.connect(lambda: self.browse_file(self.entry_rubric, "Select Rubric", "Documents (*.docx *.pdf *.txt)"))
        rubric_layout.addWidget(btn_browse_rubric)
        layout.addLayout(rubric_layout)

        # --- Assignment Knowledge Base Selection ---
        layout.addWidget(QLabel('Assignment Knowledge Base (Optional):'))
        assign_kb_layout = QHBoxLayout()
        self.entry_assign_kb = QLineEdit()
        assign_kb_layout.addWidget(self.entry_assign_kb)
        btn_browse_assign_kb = QPushButton('Browse')
        btn_browse_assign_kb.clicked.connect(lambda: self.browse_file(self.entry_assign_kb, "Select Assignment KB", "Documents (*.docx *.pdf *.txt)"))
        assign_kb_layout.addWidget(btn_browse_assign_kb)
        layout.addLayout(assign_kb_layout)

        # --- General Knowledge Base Selection ---
        layout.addWidget(QLabel('General Knowledge Base (Optional):'))
        general_kb_layout = QHBoxLayout()
        self.entry_general_kb = QLineEdit()
        general_kb_layout.addWidget(self.entry_general_kb)
        btn_browse_general_kb = QPushButton('Browse')
        btn_browse_general_kb.clicked.connect(lambda: self.browse_file(self.entry_general_kb, "Select General KB", "Documents (*.docx *.pdf *.txt)"))
        general_kb_layout.addWidget(btn_browse_general_kb)
        layout.addLayout(general_kb_layout)

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

    def browse_file(self, line_edit, title, filter_str):
        file_path, _ = QFileDialog.getOpenFileName(self, title, "", filter_str + ";;All Files (*)")
        if file_path:
            line_edit.setText(file_path)

    def run_commenter(self):
        doc_path = self.entry_doc.text().strip()
        api_key = self.entry_api_key.text().strip()
        rubric_path = self.entry_rubric.text().strip()
        assign_kb_path = self.entry_assign_kb.text().strip()
        general_kb_path = self.entry_general_kb.text().strip()

        if not api_key:
            QMessageBox.critical(self, "Error", "Please provide an OpenAI API Key.")
            return

        if not doc_path:
            QMessageBox.critical(self, "Error", "Please select an input document.")
            return
        if not os.path.exists(doc_path):
            QMessageBox.critical(self, "Error", f"Input document not found: {doc_path}")
            return

        try:
            self.lbl_status.setText("Processing with LLM...")
            QApplication.processEvents()

            # Set API key
            os.environ["OPENAI_API_KEY"] = api_key
            
            # Generate annotations via LLM
            annotations = comment.generate_annotations(
                doc_path, 
                rubric_path, 
                assign_kb_path, 
                general_kb_path
            )
            
            self.lbl_status.setText("Annotating document...")
            QApplication.processEvents()
            
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

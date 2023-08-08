##google api image to text

from google.cloud import vision
import io
from docx import Document

# Google Cloud kimlik bilgilerini ayarla
import os
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'vision.json'

# Vision API istemcisini oluştur
client = vision.ImageAnnotatorClient()

# Resmi yükle
image_path = 'tutanak.png'
with io.open(image_path, 'rb') as image_file:
    content = image_file.read()

image = vision.Image(content=content)

# Metni çıkar
response = client.text_detection(image=image)
texts = response.text_annotations

# Algılanan metni bir Word belgesine yazdır
doc = Document()
for text in texts:
    doc.add_paragraph(text.description)

# Word belgesini kaydet
doc.save('cikti.docx')

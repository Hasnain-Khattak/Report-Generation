from PIL import Image
import io

# Your binary PNG data
evidence_images = {
    'Customer Feedback Register.xlsx': b'\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x0b\xe5\x00\x00\x04&\x08\x02\x00\x00\x00c\xb9h\x0c\x00\x01\x00\x00IDATx\x9c\xec\xfdM\xc8\x1bI\x9e\xe8\xfbKE\xef\xc6\xfb\xbau\xe6R\xfeCI^\x183\xb5h\xe6\x0c\xa4\xf6s\x91\x0c\x8d\x19h\xafn\xe3]\x8a\x86...\xb1\xd6\xba{\x04\x80\x8b\xcd\xcc\xdd#\x00\x00\x00\x00\x00\x00\x00\x1c\xbf\xdf\x17I1\xa0v\xd4\xfc\x06\x9c\x00\x00\x00\x00IEND\xaeB`\x82'
}

# Convert and save the image
for file, image_bytes in evidence_images.items():
    image_stream = io.BytesIO(image_bytes)
    image = Image.open(image_stream)
    output_filename = file.replace('.xlsx', '.png')  # Example: Save as PNG with related name
    image.save(output_filename)
    print(f"Saved image as {output_filename}")

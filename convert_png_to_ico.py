from PIL import Image

# Convert PNG to ICO
img = Image.open(r"MAster/Desktopicon.png")
img.save(r"MAster/Desktopicon.ico", format="ICO", sizes=[(256,256)])
print("Icon conversion complete.")

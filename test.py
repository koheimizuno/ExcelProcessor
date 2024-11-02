import base64

# Read Excel file
with open("test.xlsx", "rb") as file:
    content = file.read()
    base64_content = base64.b64encode(content).decode()
    print(base64_content)
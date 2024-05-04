import xlwings as xw
import os

# Connect to the existing workbook
wb = xw.Book('TangoVocabSheet.xlsm')

# Get the ribbon object
ribbon = wb.api.ActiveWorkbook.Ribbon

# Create a custom ribbon tab
tab = ribbon.Tabs.Add("customTab", "Special Sauce", insertAfterMso="TabHome")

# Create a group on the tab
group = tab.Groups.Add("Scripts", "Fill It In", autoScale=True)

# Create a button group on the group
buttongroup = group.Controls.AddButtonGroup("buttongroup1")

# Load the ICO images
image_folder = r"C:\Users\dkbar\Downloads\Excel Icons"
icon_filenames = ['green-tea.ico', 'karate.ico', 'kabuki.ico']
icon_images = [os.path.join(image_folder, filename) for filename in icon_filenames]

# Check if image files exist
for img in icon_images:
    if not os.path.exists(img):
        raise FileNotFoundError(f"Icon {img} does not exist.")

# Define button attributes
buttons_info = [
    ("basic_brew_button", "Basic Brew", "RunBasicBrew", icon_images[0]),
    ("kanji_karate_button", "Kanji Karate", "RunKanjiKarate", icon_images[1]),
    ("kunyomi_kabuki_button", "Kunyomi Kabuki", "RunKunyomiKabuki", icon_images[2]),
]

# Create buttons and assign event handlers
button_handlers = {}

def create_button(buttongroup, name, label, handler_name, image_path):
    button = buttongroup.Controls.AddButton(name, label, handler_name)
    button.Image = image_path
    button.IconIndex = 0  # Use the first icon in the ICO file
    button.ShowImage = True
    button.Size = "large"
    button.Tag = handler_name.lower()
    return button

# Button click event handlers
def RunBasicBrew(control):
    print("Basic Brew button 行こう！")

def RunKanjiKarate(control):
    print("Kanji Karate button 行こう！")

def RunKunyomiKabuki(control):
    print("Kunyomi Kabuki button 行こう！")

event_handlers = {
    "RunBasicBrew": RunBasicBrew,
    "RunKanjiKarate": RunKanjiKarate,
    "RunKunyomiKabuki": RunKunyomiKabuki
}

for name, label, handler_name, image_path in buttons_info:
    button = create_button(buttongroup, name, label, handler_name, image_path)
    button.OnClick = event_handlers[handler_name]
from pyjab.jabdriver import JABDriver

app = JABDriver(
    title="Save",
    bridge_dll=r"C:\Program Files\Java\jre1.8.0_481\bin\WindowsAccessBridge-64.dll"
)

for e in app.find_elements():

    try:

        info = e.get_element_information()

        if info["role"] in [
            "label",
            "text",
            "combo box",
            "push button"
        ]:

            print(
                f"{info['role']} | "
                f"{info['name']}"
            )

    except:
        pass
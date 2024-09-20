# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno
import urllib.request
import json
import os

def get_lmstudio_output(prompt):
    """Function to get output from LMStudio"""
    lmstudio_url = 'http://localhost:1234/v1/chat/completions'  # Adjust the URL and port as needed
    
    # Create the structured payload with messages
    payload = json.dumps({
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ]
    }).encode('utf-8')

    req = urllib.request.Request(lmstudio_url, data=payload, headers={'Content-Type': 'application/json'})
    
    try:
        with urllib.request.urlopen(req) as response:
            if response.status == 200:
                response_data = response.read()
                json_response = json.loads(response_data)

                # Extracting the message content
                content = json_response['choices'][0]['message']['content']
                return content
            else:
                return 'Error retrieving output: {}'.format(response.status)
    except Exception as e:
        return 'Exception occurred: {}'.format(str(e))

def write_to_desktop(filename, content):
    """Writes content to a file on the desktop"""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", filename)
    with open(desktop_path, 'w', encoding='utf-8') as file:
        file.write(content)

def PythonApi(prompt):
    """Generates content from LMStudio and writes it to a file on the desktop"""
    # Generate content using LMStudio
    lm_output = get_lmstudio_output(prompt)

    # Write the generated content to the desktop
    filename = "lmstudio_output.txt"
    write_to_desktop(filename, lm_output)
    
    # Get the document from the scripting context which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    
    # Check whether there's already an opened document. Otherwise, create a new one
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
    
    # Get the XText interface
    text = model.Text
    
    # Create an XTextRange at the end of the document
    tRange = text.End
    
    # Set the string to the LM output
    tRange.String = lm_output
    
    return None

def create_input_dialog():
    """Creates a dialog for user input"""
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dialog_provider = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)

    # Correctly access the dialog
    try:
        dialog = dialog_provider.createDialog("vnd.sun.star.script:Standard.MyDialog?language=Basic&location=application")
    except Exception as e:
        print("Error accessing dialog in 'local':", e)
        try:
            dialog = dialog_provider.createDialog("vnd.sun.star.script:Standard.MyDialog?language=Basic&location=local")
        except Exception as e:
            print("Error accessing dialog in 'application':", e)
            dialog = dialog_provider.createDialog("vnd.sun.star.script:Standard.MyDialog?language=Basic&location=user")

    # Show the dialog
    dialog.execute()

    # Get the input from the text box
    text_box = dialog.getControl("textBox")
    prompt = text_box.getText()
    dialog.endExecute()

    return prompt


def main(*args):
    """Main function to run the dialog and process the input"""
    prompt = create_input_dialog()
    if prompt:
        PythonApi(prompt)

# Add this function call to run the main process
main()

# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno
import urllib.request
import json
import os

def get_lmstudio_output(prompt):
    """Function to get output from LMStudio"""
    lmstudio_url = 'http://localhost:1234/v1/chat/completions'  # Adjust the URL and port as needed
    
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

def PythonApi(prompt, length_option, perspective_option, voice_type):
    """Generates content from LMStudio, considering length and perspective, and writes it to a file on the desktop"""
    
    modified_prompt = f"""{perspective_option.capitalize()} perspective 
            and {length_option.lower()} text and {voice_type.lower()} voice: {prompt}"""
            
    lm_output = get_lmstudio_output(modified_prompt)
    
    filename = "lmstudio_output.txt"
    write_to_desktop(filename, lm_output)
    
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    
    if not hasattr(model, "Text"):
        model = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
    
    text = model.Text
    tRange = text.End
    tRange.String = lm_output
    
    return None

def create_input_dialog():
    ctx = uno.getComponentContext()
    smgr = ctx.getServiceManager()
    dialog_provider = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)

    dialog = dialog_provider.createDialog("vnd.sun.star.script:Standard.MyDialog?language=Basic&location=application")
    
    prompt_list = [None, None, None, None]  # TextBox, DialogType, PersonPerspective, VoiceType

    # Execute the dialog
    dialog.execute()
    
    text_box = dialog.getControl("TextBox")
    prompt_list[0] = text_box.getText()  # Get text from the input field
    
    # Get selected dialog type
    if dialog.getControl("Expanded").getModel().State == 1:
        prompt_list[1] = "Expanded"
    elif dialog.getControl("Simple").getModel().State == 1:
        prompt_list[1] = "Simple"
    elif dialog.getControl("Concise").getModel().State == 1:
        prompt_list[1] = "Concise"

    # Get selected person perspective
    if dialog.getControl("First").getModel().State == 1:
        prompt_list[2] = "First"
    elif dialog.getControl("Third").getModel().State == 1:
        prompt_list[2] = "Third Person"

    # Get selected voice type
    if dialog.getControl("Active").getModel().State == 1:
        prompt_list[3] = "Active"
    elif dialog.getControl("Passive").getModel().State == 1:
        prompt_list[3] = "Passive"

    dialog.endExecute()  # Close the dialog

    return prompt_list

def main(*args):
    """Main function to run the dialog and process the input"""
    result = create_input_dialog()  # Wait for input
    
    # Unpack all four elements from result
    prompt, length_option, perspective_option, voice_type = result
    
    if prompt and length_option and perspective_option and voice_type:  # If inputs are not empty, process them
        PythonApi(prompt, length_option, perspective_option, voice_type)

# Call the main process
main()

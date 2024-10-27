import uno
import urllib.request
import json

def show_input_dialog():
    """Affiche une boîte de dialogue pour que l'utilisateur saisisse son besoin."""
    ctx = uno.getComponentContext()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)

    doc = desktop.getCurrentComponent()

    # Création de la boîte de dialogue
    dialog = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialog", ctx)
    dialog_model = smgr.createInstanceWithContext("com.sun.star.awt.UnoControlDialogModel", ctx)
    
    dialog.setModel(dialog_model)
    dialog.setVisible(False)
    dialog_model.Title = "Saisissez votre besoin"
    dialog_model.Width = 400  # Augmentation de la largeur
    dialog_model.Height = 200  # Augmentation de la hauteur

    # Ajout d'une grande zone de texte pour le besoin utilisateur
    edit_model = dialog_model.createInstance("com.sun.star.awt.UnoControlEditModel")
    edit_model.Width = 380  # Largeur du champ de texte
    edit_model.Height = 100  # Hauteur du champ de texte pour plus d'espace d'écriture
    edit_model.MultiLine = True  # Autoriser le texte sur plusieurs lignes
    edit_model.PositionX = 10
    edit_model.PositionY = 10
    edit_model.Name = "user_input"
    dialog_model.insertByName("UserInputEdit", edit_model)

    # Ajout d'un bouton OK
    button_model = dialog_model.createInstance("com.sun.star.awt.UnoControlButtonModel")
    button_model.Width = 50
    button_model.Height = 20
    button_model.PositionX = 170
    button_model.PositionY = 130
    button_model.Name = "OKButton"
    button_model.Label = "OK"
    dialog_model.insertByName("OKButton", button_model)

    dialog.setVisible(True)
    dialog.execute()

    # Récupération de l'entrée utilisateur
    user_input = dialog.getControl("UserInputEdit").getText()
    dialog.endExecute()

    if not user_input:
        return None
    
    return user_input

def connect_to_llm(prompt, ip_address="127.0.0.1", port=1234):
    """Envoie une requête au LLM pour générer du texte à partir du prompt fourni."""
    url = f"http://{ip_address}:{port}/v1/completions"
    data = json.dumps({
        "prompt": prompt,
        "max_tokens": 300  # Augmenter la limite des tokens pour obtenir des réponses plus longues
    }).encode('utf-8')

    req = urllib.request.Request(url, data=data, headers={'Content-Type': 'application/json'})
    
    try:
        with urllib.request.urlopen(req) as response:
            if response.status == 200:
                result = json.loads(response.read().decode('utf-8'))
                return result['choices'][0]['text']
            else:
                return f"Erreur : {response.status}"
    except urllib.error.URLError as e:
        return f"Erreur de connexion : {e}"

def insert_text(document, text):
    """Insère du texte dans un document LibreOffice Writer."""
    cursor = document.Text.createTextCursor()
    document.Text.insertString(cursor, text, 0)

def main():
    # Affichage de la boîte de dialogue pour saisir le besoin
    user_input = show_input_dialog()

    if user_input:
        # Se connecter au LLM et générer le texte basé sur le besoin utilisateur
        generated_text = connect_to_llm(user_input)
        
        # Se connecter à LibreOffice et insérer le texte généré dans un document Writer
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        context = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        desktop = context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", context)
        document = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        
        insert_text(document, generated_text)
    else:
        print("Aucun besoin saisi, annulation.")

# Exécuter la fonction principale
if __name__ == "__main__":
    main()

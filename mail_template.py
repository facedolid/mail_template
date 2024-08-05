import PySimpleGUI as sg
import win32com.client as win32
import json
import os

# ユーティリティ関数
def load_json(filename, default=None):
    if default is None:
        default = {}
    if os.path.exists(filename):
        try:
            with open(filename, "r", encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            print(f"エラー: {filename} の解析に失敗しました。")
            return default
    return default

def save_json(filename, data):
    try:
        with open(filename, "w", encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"エラー: {filename} の保存に失敗しました。{str(e)}")

# Outlook関連の関数
def create_outlook_email(to, cc, bcc, subject, body, signature, attachment_paths=None):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.CC = cc
        mail.BCC = bcc
        mail.Subject = subject
        mail.Body = body + "\n\n" + signature
        
        if attachment_paths:
            for path in attachment_paths:
                if os.path.exists(path):
                    mail.Attachments.Add(path)
        
        mail.Display(True)
    except Exception as e:
        sg.popup_error(f"メール作成エラー: {str(e)}")

# プリセット関連の関数
def load_presets():
    return load_json("presets.json")

def save_preset(name, to, cc, bcc, subject, body):
    presets = load_presets()
    presets[name] = {"to": to, "cc": cc, "bcc": bcc, "subject": subject, "body": body}
    save_json("presets.json", presets)

def delete_preset(name):
    presets = load_presets()
    if name in presets:
        del presets[name]
        save_json("presets.json", presets)

# 署名関連の関数
def load_signatures():
    return load_json("signatures.json")

def save_signature(name, content):
    signatures = load_signatures()
    signatures[name] = content
    save_json("signatures.json", signatures)

def delete_signature(name):
    signatures = load_signatures()
    if name in signatures:
        del signatures[name]
        save_json("signatures.json", signatures)

def get_default_signature(signatures):
    return 'default' if 'default' in signatures else (list(signatures.keys())[0] if signatures else '')

# 連絡先関連の関数
def load_contacts():
    return load_json("contacts.json")

def save_contact(name, email):
    contacts = load_contacts()
    contacts[name] = email
    save_json("contacts.json", contacts)

def delete_contact(name):
    contacts = load_contacts()
    if name in contacts:
        del contacts[name]
        save_json("contacts.json", contacts)

def create_main_layout():
    presets = load_presets()
    signatures = load_signatures()
    contacts = load_contacts()
    
    preset_names = list(presets.keys())
    signature_names = list(signatures.keys())
    contact_names = list(contacts.keys())
    
    default_signature = get_default_signature(signatures)

    layout = [
        [sg.Frame('プリセット', [
            [sg.Combo(preset_names, key='-PRESET-', enable_events=True, size=(30, 1)),
             sg.Button('適用'), sg.Button('管理')]
        ])],
        [sg.Frame('メール作成', [
            [sg.Text('宛先:'), sg.Input(key='-TO-', size=(30, 1)), 
             sg.Combo(contact_names, key='-TO-CONTACT-', enable_events=True, size=(20, 1))],
            [sg.Text('CC:'), sg.Input(key='-CC-', size=(30, 1)), 
             sg.Combo(contact_names, key='-CC-CONTACT-', enable_events=True, size=(20, 1))],
            [sg.Text('BCC:'), sg.Input(key='-BCC-', size=(30, 1)), 
             sg.Combo(contact_names, key='-BCC-CONTACT-', enable_events=True, size=(20, 1))],
            [sg.Text('件名:'), sg.Input(key='-SUBJECT-', size=(50, 1))],
            [sg.Text('本文:')],
            [sg.Multiline(size=(60, 10), key='-BODY-')],
            [sg.Text('署名:'), sg.Combo(signature_names, default_value=default_signature, key='-SIGNATURE-', enable_events=True, size=(20, 1))],
            [sg.Text('添付ファイル:')],
            [sg.Listbox(values=[], size=(50, 3), key='-ATTACHMENT-LIST-')],
            [sg.Input(key='-FILE-', visible=False, enable_events=True), 
             sg.FilesBrowse('ファイル追加'), sg.Button('選択したファイルを削除')]
        ])],
        [sg.Button('メール作成'), sg.Button('署名管理'), sg.Button('連絡先管理'), sg.Button('終了')]
    ]
    return layout

def create_preset_management_window():
    presets = load_presets()
    preset_names = list(presets.keys())
    layout = [
        [sg.Listbox(values=preset_names, size=(30, 6), key='-PRESET-LIST-')],
        [sg.Button('新規'), sg.Button('編集'), sg.Button('削除'), sg.Button('閉じる')]
    ]
    return sg.Window('プリセット管理', layout)

def create_signature_management_window():
    signatures = load_signatures()
    signature_names = list(signatures.keys())
    layout = [
        [sg.Listbox(values=signature_names, size=(30, 6), key='-SIGNATURE-LIST-')],
        [sg.Button('新規'), sg.Button('編集'), sg.Button('削除'), sg.Button('閉じる')]
    ]
    return sg.Window('署名管理', layout)

def create_contact_management_window():
    contacts = load_contacts()
    contact_names = list(contacts.keys())
    layout = [
        [sg.Listbox(values=contact_names, size=(30, 6), key='-CONTACT-LIST-')],
        [sg.Button('新規'), sg.Button('編集'), sg.Button('削除'), sg.Button('閉じる')]
    ]
    return sg.Window('連絡先管理', layout)

def edit_preset_window(preset_name, preset_data):
    layout = [
        [sg.Text('プリセット名:'), sg.Input(preset_name, key='-NAME-')],
        [sg.Text('宛先:'), sg.Input(preset_data.get('to', ''), key='-TO-')],
        [sg.Text('CC:'), sg.Input(preset_data.get('cc', ''), key='-CC-')],
        [sg.Text('BCC:'), sg.Input(preset_data.get('bcc', ''), key='-BCC-')],
        [sg.Text('件名:'), sg.Input(preset_data.get('subject', ''), key='-SUBJECT-')],
        [sg.Text('本文:')],
        [sg.Multiline(preset_data.get('body', ''), size=(60, 10), key='-BODY-')],
        [sg.Button('保存'), sg.Button('キャンセル')]
    ]
    window = sg.Window(f'プリセット編集: {preset_name}', layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, 'キャンセル'):
            window.close()
            return None
        if event == '保存':
            window.close()
            return {
                "name": values['-NAME-'],
                "to": values['-TO-'],
                "cc": values['-CC-'],
                "bcc": values['-BCC-'],
                "subject": values['-SUBJECT-'],
                "body": values['-BODY-']
            }

def edit_signature_window(signature_name, signature_content):
    layout = [
        [sg.Text('署名名:'), sg.Input(signature_name, key='-NAME-')],
        [sg.Text('署名:')],
        [sg.Multiline(signature_content, size=(60, 10), key='-SIGNATURE-')],
        [sg.Button('保存'), sg.Button('キャンセル')]
    ]
    window = sg.Window(f'署名編集: {signature_name}', layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, 'キャンセル'):
            window.close()
            return None
        if event == '保存':
            window.close()
            return {
                "name": values['-NAME-'],
                "content": values['-SIGNATURE-']
            }

def edit_contact_window(contact_name, contact_email):
    layout = [
        [sg.Text('名前:'), sg.Input(contact_name, key='-NAME-')],
        [sg.Text('メールアドレス:'), sg.Input(contact_email, key='-EMAIL-')],
        [sg.Button('保存'), sg.Button('キャンセル')]
    ]
    window = sg.Window(f'連絡先編集: {contact_name}', layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, 'キャンセル'):
            window.close()
            return None
        if event == '保存':
            window.close()
            return {
                "name": values['-NAME-'],
                "email": values['-EMAIL-']
            }

def main():
    sg.theme('LightBlue2')
    window = sg.Window('Outlook メール作成', create_main_layout(), finalize=True)
    
    attachment_list = []

    while True:
        try:
            event, values = window.read()

            if event == sg.WINDOW_CLOSED or event == '終了':
                break
            elif event == 'メール作成':
                signature = load_signatures().get(values['-SIGNATURE-'], '')
                create_outlook_email(values['-TO-'], values['-CC-'], values['-BCC-'], values['-SUBJECT-'], values['-BODY-'], signature, attachment_list)
            elif event == '適用':
                selected_preset = load_presets().get(values['-PRESET-'])
                if selected_preset:
                    window['-TO-'].update(selected_preset.get('to', ''))
                    window['-CC-'].update(selected_preset.get('cc', ''))
                    window['-BCC-'].update(selected_preset.get('bcc', ''))
                    window['-SUBJECT-'].update(selected_preset.get('subject', ''))
                    window['-BODY-'].update(selected_preset.get('body', ''))
            elif event == '管理':
                manage_presets(window)
            elif event == '署名管理':
                manage_signatures(window)
            elif event == '連絡先管理':
                manage_contacts(window)
            elif event == '-FILE-':
                new_files = values['-FILE-'].split(';')
                attachment_list.extend(new_files)
                window['-ATTACHMENT-LIST-'].update(values=attachment_list)
            elif event == '選択したファイルを削除':
                selected_files = values['-ATTACHMENT-LIST-']
                for file in selected_files:
                    attachment_list.remove(file)
                window['-ATTACHMENT-LIST-'].update(values=attachment_list)
            elif event in ('-TO-CONTACT-', '-CC-CONTACT-', '-BCC-CONTACT-'):
                field = event.split('-')[1]
                selected_contact = values[event]
                if selected_contact:
                    contacts = load_contacts()
                    current_value = values[f'-{field}-']
                    if current_value:
                        current_value += '; '
                    current_value += contacts[selected_contact]
                    window[f'-{field}-'].update(current_value)
        except Exception as e:
            sg.popup_error(f"エラーが発生しました: {str(e)}")

    window.close()

def manage_presets(main_window):
    preset_window = create_preset_management_window()
    while True:
        try:
            event, values = preset_window.read()
            if event in (sg.WINDOW_CLOSED, '閉じる'):
                break
            elif event == '新規':
                new_preset = edit_preset_window('新規プリセット', {})
                if new_preset:
                    save_preset(new_preset['name'], new_preset['to'], new_preset['cc'], new_preset['bcc'], new_preset['subject'], new_preset['body'])
                    preset_window['-PRESET-LIST-'].update(values=list(load_presets().keys()))
            elif event == '編集':
                selected_preset = values['-PRESET-LIST-'][0] if values['-PRESET-LIST-'] else None
                if selected_preset:
                    presets = load_presets()
                    edited_preset = edit_preset_window(selected_preset, presets[selected_preset])
                    if edited_preset:
                        if edited_preset['name'] != selected_preset:
                            delete_preset(selected_preset)
                        save_preset(edited_preset['name'], edited_preset['to'], edited_preset['cc'], edited_preset['bcc'], edited_preset['subject'], edited_preset['body'])
                        preset_window['-PRESET-LIST-'].update(values=list(load_presets().keys()))
            elif event == '削除':
                selected_preset = values['-PRESET-LIST-'][0] if values['-PRESET-LIST-'] else None
                if selected_preset:
                    if sg.popup_yes_no(f'プリセット "{selected_preset}" を削除しますか？', title='確認') == 'Yes':
                        delete_preset(selected_preset)
                        preset_window['-PRESET-LIST-'].update(values=list(load_presets().keys()))
        except Exception as e:
            sg.popup_error(f"エラーが発生しました: {str(e)}")
    preset_window.close()
    main_window['-PRESET-'].update(values=list(load_presets().keys()))

def manage_signatures(main_window):
    signature_window = create_signature_management_window()
    while True:
        try:
            event, values = signature_window.read()
            if event in (sg.WINDOW_CLOSED, '閉じる'):
                break
            elif event == '新規':
                new_signature = edit_signature_window('新規署名', '')
                if new_signature:
                    save_signature(new_signature['name'], new_signature['content'])
                    signature_window['-SIGNATURE-LIST-'].update(values=list(load_signatures().keys()))
            elif event == '編集':
                selected_signature = values['-SIGNATURE-LIST-'][0] if values['-SIGNATURE-LIST-'] else None
                if selected_signature:
                    signatures = load_signatures()
                    edited_signature = edit_signature_window(selected_signature, signatures[selected_signature])
                    if edited_signature:
                        if edited_signature['name'] != selected_signature:
                            delete_signature(selected_signature)
                        save_signature(edited_signature['name'], edited_signature['content'])
                        signature_window['-SIGNATURE-LIST-'].update(values=list(load_signatures().keys()))
            elif event == '削除':
                selected_signature = values['-SIGNATURE-LIST-'][0] if values['-SIGNATURE-LIST-'] else None
                if selected_signature:
                    if sg.popup_yes_no(f'署名 "{selected_signature}" を削除しますか？', title='確認') == 'Yes':
                        delete_signature(selected_signature)
                        signature_window['-SIGNATURE-LIST-'].update(values=list(load_signatures().keys()))
        except Exception as e:
            sg.popup_error(f"エラーが発生しました: {str(e)}")
    signature_window.close()
    signatures = load_signatures()
    main_window['-SIGNATURE-'].update(values=list(signatures.keys()), value=get_default_signature(signatures))

def manage_contacts(main_window):
    contact_window = create_contact_management_window()
    while True:
        try:
            event, values = contact_window.read()
            if event in (sg.WINDOW_CLOSED, '閉じる'):
                break
            elif event == '新規':
                new_contact = edit_contact_window('新規連絡先', '')
                if new_contact:
                    save_contact(new_contact['name'], new_contact['email'])
                    contact_window['-CONTACT-LIST-'].update(values=list(load_contacts().keys()))
            elif event == '編集':
                selected_contact = values['-CONTACT-LIST-'][0] if values['-CONTACT-LIST-'] else None
                if selected_contact:
                    contacts = load_contacts()
                    edited_contact = edit_contact_window(selected_contact, contacts[selected_contact])
                    if edited_contact:
                        if edited_contact['name'] != selected_contact:
                            delete_contact(selected_contact)
                        save_contact(edited_contact['name'], edited_contact['email'])
                        contact_window['-CONTACT-LIST-'].update(values=list(load_contacts().keys()))
            elif event == '削除':
                selected_contact = values['-CONTACT-LIST-'][0] if values['-CONTACT-LIST-'] else None
                if selected_contact:
                    if sg.popup_yes_no(f'連絡先 "{selected_contact}" を削除しますか？', title='確認') == 'Yes':
                        delete_contact(selected_contact)
                        contact_window['-CONTACT-LIST-'].update(values=list(load_contacts().keys()))
        except Exception as e:
            sg.popup_error(f"エラーが発生しました: {str(e)}")
    contact_window.close()
    contacts = load_contacts()
    main_window['-TO-CONTACT-'].update(values=list(contacts.keys()))
    main_window['-CC-CONTACT-'].update(values=list(contacts.keys()))
    main_window['-BCC-CONTACT-'].update(values=list(contacts.keys()))

if __name__ == '__main__':
    main()
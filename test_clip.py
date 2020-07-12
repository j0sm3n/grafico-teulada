# import subprocess


# def get_clipboard():
#     p = subprocess.Popen(
#         ['xclip', '-selection', 'clipbard', '-o'],
#         stdout=subprocess.PIPE)
#     retcode = p.wait()
#     data = p.stdout.read()
#     return data


# text = str(get_clipboard())
# text_list = text.split('\n')

# print(text)

# for t in text_list:
#     print(t)

# text = get_clipboard()
# print(type(text))



import pyperclip

# pyperclip.copy('Te text to be copied to the clipboard.')

text = pyperclip.paste()
text.replace('\n', '')
text_list = text.split()
print(type(text))
print(text)
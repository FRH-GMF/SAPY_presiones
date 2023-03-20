# El codigo siguiente es utilizado para codificar las imagenes e iconos que se encuentran
# en formato bytes a ASCII de manera de introducirlo dentro del codigo sin necesidad de
# tener que lidiar con archivos externos cuando se compila el programa en .exe.
import base64


def pic2str(file, functionName):
    pic = open(file, 'rb')
    content = '{} = {}\n'.format(functionName, base64.b64encode(pic.read()))
    pic.close()

    with open('converted_string.py', 'a') as f:
        f.write(content)


if __name__ == '__main__':
    pic2str('Logo LAyF.png', 'string')

# import base64
# logo = open('Logo LAyF.png', 'rb')
# encoded = base64.b64encode(logo.read())
# 'ZGF0YSB0byBiZSBlbmNvZGVk'
# data = base64.b64decode(encoded)
#
# with open('logo.py', 'a') as f:
#     f.write(str(encoded))


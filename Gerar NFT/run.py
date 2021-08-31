from PIL import Image
from random import randint

imagensGeradas = []

def gerarImagem(nomeImagem):
    dados = []
    dados.append(str(randint(0,4)))
    dados.append(str(randint(0,5)))
    dados.append(str(randint(0,5)))

    valida = dados[0] + "|" + dados[1] + "|" + dados[2]

    if valida in imagensGeradas:
      print(str(nomeImagem) + " Imagem com a combinação (" + valida + ") já existe")
      gerarImagem(nomeImagem)
    else:
      background = Image.open("assets/backgrounds/" + dados[0] + ".png")
      skin =       Image.open("assets/skin/" + dados[1] + ".png")
      cosmetico =  Image.open("assets/cosmetico/" + dados[2] + ".png")

      background.paste(skin, (0, 0), skin)
      background.paste(cosmetico, (0, 0), cosmetico)
      background.save("image/" + str(nomeImagem) + ".png")

      imagensGeradas.append(valida)

for x in range(50):
  gerarImagem(x)

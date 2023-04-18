import time

from pynput.mouse import Button, Controller
from time import sleep

mouse = Controller()
#mouse.position = (15,756) # Bota o mouse nessa posição e fala a posição do mouse quando não tem nada dentro
#mouse.move(0,-100) # Move o mouse um determinado número de pixels nos eixo x e y
#mouse.click(Button.left,2) # Faz o mouse clicar, quantas vezes quiser
#mouse.press(Button.left) # Faz o mouse clicar e segurar, matém clicado/pressionado
#mouse.release(Button.left) # Faz o mouse soltar o botão
#mouse.scroll(-100, 0) # Faz o mouse "scrollar" no eixo x ou y

# mouse.position = (1304,40)
# mouse.click(Button.left, 2)
# mouse.position = (683, 352)
# time.sleep(1)
# mouse.click(Button.left, 2)

mouse.press(Button.left)
mouse.move(0,100)
time.sleep(1)
mouse.release(Button.left)
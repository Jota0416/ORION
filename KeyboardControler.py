from pynput.keyboard import Key, Controller

keyb = Controller()

# keyb.press('1') # Pressiona alguma tecla alfanumérica
# keyb.release('1') # Sempre usar o release depois do press

# keyb.type('É o flamengo krl!') # Usado para escrever texto

# keyb.press(Key.cmd) #Pras teclas mais diferentes tem que usar esse "Key."
# keyb.release(Key.cmd)

with keyb.pressed(Key.ctrl, Key.alt):
    keyb.press('l')
    keyb.release('l')

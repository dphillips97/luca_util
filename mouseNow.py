#! python3
# mouseNow.py - Displays the mouse cursor's current position.
import pyautogui as pag
print('Press CTRL-C to quit')
try:
	while True:
		x, y = pag.position()
		
		positionStr = 'x: ' + str(x).rjust(4) + ' y: ' + str(y).rjust(4) 
		
		pixelColor = pag.screenshot().getpixel((x,y))
		
		positionStr += ' RGB: (' + str(pixelColor[0]).rjust(3)
		positionStr += ', ' + str(pixelColor[1]).rjust(3)
		positionStr += ', ' + str(pixelColor[2]).rjust(3) + ')'
		
		
		print(positionStr, end = '')
		
		print('\b' * len(positionStr), end = '', flush = True)
			
		
		
except KeyboardInterrupt:
	print('\nDone.')
		
		
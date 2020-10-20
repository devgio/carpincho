from datetime import datetime
from colorama import init, Fore, Back, Style
init(autoreset=True)

ur, ll, ul, lr, ve, ho = '╗', '╚', '╔', '╝', '║', '═'
blue_white_style = Back.WHITE+Fore.BLUE+Style.BRIGHT

def header(title):
	divider = ho * (80)
	def hashedRow(text):
		s = f'{ve} {text} '
		space = ' '*(len(divider) - len(s) + 1)
		print(f'{blue_white_style}{s}{space}{ve}')
	print(f'{blue_white_style}{ul}{divider}{ur}')
	hashedRow(title)
	hashedRow('')
	runtime = datetime.now().strftime('%d %B %Y @ %H:%M')
	hashedRow(f'Runtime: {runtime}')
	print(f'{blue_white_style}{ll}{divider}{lr}\n')

def timed_row(text, tag=False, tag_style='', keep_row=False):
	time_stamp = datetime.now().strftime('%H:%M')
	line = f'[{Fore.CYAN}{time_stamp}{Style.RESET_ALL}]'
	if tag:
		if tag_style:
			tag_style_dict = {
				'error':f'{Back.RED}{Style.BRIGHT}{Fore.WHITE}', 
				'warning':f'{Back.YELLOW}{Fore.BLACK}', 
				'ok':f'{Style.BRIGHT}{Back.GREEN}', 
				'attention':f'{Style.BRIGHT}{Fore.MAGENTA}'
				}
			line += f'[{tag_style_dict[tag_style.lower()]}{tag}{Style.RESET_ALL}]'
	line += f' {text}'
	if keep_row:
		print(line, end='\r', flush=True)
		return
	print(line)






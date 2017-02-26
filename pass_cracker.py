########################################################################################
### Pass Cracker
### Description: Word Documents password cracker based on dictionary attack
### Author: Daniel Fernandez Perez
### Notes: 
###  - Based on the code by Gabe Marshall: https://gist.github.com/gabemarshall/9372073
###  - sys.exit(1) if finished without errors
###  - sys.exit(2) if finished with any error
########################################################################################

# Required imports:
## sys for accessing command line arguments and system functions
## getopt for parsing coomand line arguments
## win32.com for using windows applications, external library, requires instalation

import sys, getopt, win32com.client

# displayHelp - Function to display usage help
# input:
## none
# output:
## usage help message on console
		
def displayHelp():
	print ('Usage: pass_cracker.py -w <wordlist> -f <file>')
	print ('When specifying the file, enter full path.')
	
# main - Function to handle the main functionality of the program
# input:
## argv, list or arguments user passed to the program
# output:
## success or error message, help information

def main(argv):
	wordlist = ''
	filename = ''
	found = False
	
	try:
		word = win32com.client.Dispatch("Word.Application")
	except:
		print('Unable to use win32com library, please check instalation')
		sys.exit(2)
	
	try:
		opts, args = getopt.getopt(argv,"hw:f:",["wordlist=","file="])
	except getopt.GetoptError:
		displayHelp()
		sys.exit(2)
	if not argv:
		print('No arguments have been passed.')
		displayHelp()
		sys.exit(2)
	elif len(argv)<4:
		print('Invalid arguments passed.')
		displayHelp()
		sys.exit(2)
	else:
		for opt, arg in opts:
			if opt.lower() in ("-?","-h","-help"):
				displayHelp()
				sys.exit(1)
			elif opt.lower() in ("-w","--wordlist"):
				wordlist = arg.lower()
			elif opt.lower() in ("-f","--file"):
				filename = arg.lower()
	
	try:
		open(filename, 'r')
	except:
		print('Cannot open file')
		sys.exit(2)
		
	try:
		password_file = open(wordlist, 'r')
		passwords = password_file.readlines()
		password_file.close()
		passwords = [item.rstrip('\n') for item in passwords]
	except:
		print('Cannot open wordlist')
		sys.exit(2)
		
	try:
		results = open('results.txt', 'w')
	except:
		print('Cannot open/create results doc')
		sys.exit(2)
		
	for password in passwords:
		if (found == False):
			#print(password)
			try:
				wb = word.Documents.Open(filename, False, True, None, password)
				if (word.Documents[0].Content):
					found = True
					print('Password found: ' + password)
					print('Password has been stored in results.txt file.')
					results.write(password)
					results.close()
			except:
				#no output to be displayed
				pass
		else:
			break
	sys.exit(1)
	
if __name__ == "__main__":
	main(sys.argv[1:])
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

import sys, getopt, time, win32com.client

# displayHelp - Function to display usage help
# input:
## none
# output:
## usage help message on console
		
def displayHelp():
	print ('Usage: pass_cracker.py -w <wordlist> -f <file>')
	print ('When specifying the file, enter full path.')

# timeit - Function to calculate the time a given function takes
# input
## funtion to monitor
# output
## message on console with time function took to execute
# Note: function adapted from 
# http://codereview.stackexchange.com/questions/151928/random-password-cracker-using-brute-force

def timeit(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        print ('The password has been found in {:.5f} seconds.'.format(time.time() - start))
        return result
    return wrapper

# getPass - Function get password from file with a given dictionary	
# input:
## word, object with MS Word functions
## filename, file to open
## passwords, all passwords from wordlist file
## results,file to store password
# output:
## if password is found, it is stored in results file. Message on console.

@timeit
def getPass(word, filename, passwords, results):
	found = False
	
	for password in passwords:
		if (found == False):
			try:
				word.Documents.Open(filename, False, True, None, password)
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
	
# main - Function to handle the main functionality of the program
# input:
## argv, list or arguments user passed to the program
# output:
## success or error message, help information

def main(argv):
	wordlist = ''
	filename = ''
	
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
				wordlist = arg
			elif opt.lower() in ("-f","--file"):
				filename = arg
	
	try:
		open(filename, 'r')
	except:
		print('Cannot access file provided.')
		sys.exit(2)
		
	try:
		password_file = open(wordlist, 'r')
		passwords = password_file.readlines()
		password_file.close()
		passwords = [item.rstrip('\n') for item in passwords]
	except:
		print('Cannot open wordlist.')
		sys.exit(2)
		
	try:
		results = open('results.txt', 'w')
	except:
		print('Cannot create/modify results doc.')
		sys.exit(2)
		
	print('Trying to open file: ' + '"' + filename + '"')
	print('With dictionary: ' +  '"' + wordlist + '"')

	getPass(word, filename, passwords, results)
	
	sys.exit(1)
	
if __name__ == "__main__":
	# First argument is the name of the program, so it can be ignored
	main(sys.argv[1:])
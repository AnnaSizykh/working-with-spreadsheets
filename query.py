import functions

def query(ws):
    print('description')
    while True:
        query = input('query: ')
        if not query:
            print('query is empty.')
        elif query == 'stop':
            break
        elif query == 'qhelp':
            print('help field')
        elif query == 'add':
            data = input('Addition data description\n').split()
            functions.addition(ws, data[0], data[1], data[2])
        elif query == 'sub':
            data = input('subtraction data description\n').split()
            functions.subtraction(ws, data[0], data[1], data[2])
        elif query == 'multi':
            data = input('multiplication data description\n').split()
            functions.multiplication(ws, data[0], data[1], data[2])
        elif query == 'div':
            data = input('division data description\n').split()
            functions.division(ws, data[0], data[1], data[2])
        elif query == 'rounded':
            data = input('round data description\n').split()
            functions.round_number(ws, data[0], int(data[1]), data[2])
        elif query == 'expo':
            data = input('exponentiation data description\n').split()
            functions.exponentiation(ws, data[0], int(data[1]), data[2])
        elif query == 'log':
            data = input('logarithm data description\n').split()
            functions.logarithm(ws, data[0], int(data[1]), data[2])
        elif query == 'mean':
            data = input('mean data description\n').split()
            functions.mean(ws, data[0], data[1], data[2])
        elif query == 'move':
            data = input('move data description\n').split()
            functions.move(ws, data[0], data[1], data[2], data[3])
        elif query == 'copy':
            data = input('copy data description\n').split()
            functions.copy(ws, data[0], data[1], data[2], data[3])
        elif query == 'delete':
            data = input('delete data description\n').split()
            functions.delete(ws, data[0], data[1])
        elif query == 'compare':
            data = input('compare data description\n').split()
            functions.compare(ws, data[0], data[1], data[2])
        elif query == 'find':
            data = input('find data description\n')
            print(functions.find(ws, str(data)))
        else:
            print(f'function "{query}" not found. Please, try again or print "qhelp" to see available functions.')

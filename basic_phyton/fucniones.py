#Funciones

def miFuncion():
	dato = 0
	while (dato < 10):
		print(dato)
		dato+=1
        
def myFun(param = "algo"):
    print(param)
    
#funciones lambda o anonimas
# nombre_funcion = lambda param1, param2, etc : param1 * param2
sumar = lambda dato1, dato2 : dato1 * dato2

#------	
miFuncion()
myFun()

print(sumar(12,14))
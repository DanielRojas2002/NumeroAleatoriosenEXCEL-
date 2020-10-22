import openpyxl
import random

class datos:
    def __init__(self,cuantos,rango,nombre):
        self.__cuantos=cuantos
        self.__rango=rango
        self.__nombre=nombre
        
    
    def creacion(self):
        libro=openpyxl.Workbook()
        hoja=libro["Sheet"]
        hoja.title="NumerosRandom"
        hoja["B1"]="Numero Aleatorios"
        
        for renglon in range(2,self.__cuantos+2):
            hoja.cell(row=renglon,column=2).value= random.randrange(self.__rango)
            
        print("Menu de operaciones ")
        FUNCION="=CONTAR"
        print("1=Sumar\n2=Promedio\n3=ValorMaximo\n4=ValorMinimo\n5=TODAS\n6=No quiero ninguna operacion extra")
        opcion=int(input(": "))
        if opcion==1:
            string=str(self.__cuantos+1)
            funcion="=SUM"
            funcion2="(B2:"
            funcion3="B"
            funcion4=")"
            fin=funcion+funcion2+funcion3+string+funcion4
            hoja["D9"].value="Suma"
            hoja["D10"].value=fin
            hoja["D5"].value="Registros:"
            fin_1=FUNCION+funcion2+funcion3+string+funcion4
            hoja["E5"].value=fin_1
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO ")
            print("El excel esta en el directorio donde esta guardado el codigo :)")
        
        elif opcion==2:
            string=str(self.__cuantos+1)
            funcion="=PROMEDIO"
            funcion2="(B2:"
            funcion3="B"
            funcion4=")"
            fin=funcion+funcion2+funcion3+string+funcion4
            hoja["E9"].value="Promedio"
            hoja["E10"].value=fin
            
            hoja["D5"].value="Registros:"
            fin_1=FUNCION+funcion2+funcion3+string+funcion4
            hoja["E5"].value=fin_1
            
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO")
            print("El excel esta en el directorio donde esta guardado el codigo :)")
        
        elif opcion==3:
            string=str(self.__cuantos+1)
            funcion="=MAX"
            funcion2="(B2:"
            funcion3="B"
            funcion4=")"
            fin=funcion+funcion2+funcion3+string+funcion4
            hoja["F9"].value="VMaximo"
            hoja["F10"].value=fin
            hoja["D5"].value="Registros:"
            fin_1=FUNCION+funcion2+funcion3+string+funcion4
            hoja["E5"].value=fin_1
            
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO")
            print("El excel esta en el directorio donde esta guardado el codigo :)")
        
        
        elif opcion==4:
            string=str(self.__cuantos+1)
            funcion="=MIN"
            funcion2="(B2:"
            funcion3="B"
            funcion4=")"
            fin=funcion+funcion2+funcion3+string+funcion4
            hoja["G9"].value="VMinimo"
            hoja["G10"].value=fin
            hoja["D5"].value="Registros:"
            fin_1=FUNCION+funcion2+funcion3+string+funcion4
            hoja["E5"].value=fin_1
            
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO")
            print("El excel esta en el directorio donde esta guardado el codigo :)")
            
            
                
        elif opcion==5:
            string=str(self.__cuantos+1)
            funcion="=SUM"
            funcion2="(B2:"
            funcion3="B"
            funcion4=")"
            fin=funcion+funcion2+funcion3+string+funcion4
            hoja["D9"].value="Suma"
            hoja["D10"].value=fin
            
            funcion5="=PROMEDIO"
            fina=funcion5+funcion2+funcion3+string+funcion4
            hoja["E9"].value="Promedio"
            hoja["E10"].value=fina
            
            funcion6="=MAX"
            final=funcion6+funcion2+funcion3+string+funcion4
            hoja["F9"].value="VMaximo"
            hoja["F10"].value=final
            
            
            funcion7="=MIN"
            finn=funcion7+funcion2+funcion3+string+funcion4
            hoja["G9"].value="VMinimo"
            hoja["G10"].value=finn
            
            hoja["D5"].value="Registros:"
            fin_1=FUNCION+funcion2+funcion3+string+funcion4
            hoja["E5"].value=fin_1
            
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO")
            print("El excel esta en el directorio donde esta guardado el codigo :)")

        
        elif opcion==6:
            nombre2=".xlsx"
            nombref=self.__nombre+nombre2
            libro.save(nombref)
            print("EXITOSO")
            print("El excel esta en el directorio donde esta guardado el codigo :)")
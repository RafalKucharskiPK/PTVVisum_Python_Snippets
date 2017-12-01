 
 
#Visum.MatrixEditor.Gravitate("C:\Users\RK\Desktop\Szamerica\distances.dis","C:\Users\RK\Desktop\Szamerica\mat.fma" , "C:\Users\RK\Desktop\Szamerica\cod.cod", "$V;d2" )


# tutaj musisz zdefinowac swoje sciezki dla plikow cod dis i rezulatatu dzialania #

distance_mtx_path="C:\Users\RK\Desktop\Szamerica\distances.dis"
result_mtx_path="C:\Users\RK\Desktop\Szamerica\mat.fma" # ! ! ! ! ! ta macierz musi wczesniej istniec! ! !  jesli nie ma takiego pliku wyskakuje blad - mozna stworzyc pupsty plik o tej nazwie
code_file_path = "C:\Users\RK\Desktop\Szamerica\cod.cod" # ! ! ! ! w sciezce nie moze byc spacji, wiec musisz inaczej nazywac pliki
format = "$V;d2"


#procedura do stworzenia pliku .fma jesli go nie ma
try: f=open("C:\Users\RK\Desktop\Szamerica\mat.fma","r")
except:
    f=open("C:\Users\RK\Desktop\Szamerica\mat.fma","w")
    f.close()


Visum.MatrixEditor.Gravitate(distance_mtx_path,result_mtx_path,code_file_path,format)

# i teraz mozesz wykonac to drugi raz z innym .cod:

code_file_path = "C:\Users\RK\Desktop\Szamerica\cod.cod"
Visum.MatrixEditor.Gravitate(distance_mtx_path,result_mtx_path,code_file_path,format)





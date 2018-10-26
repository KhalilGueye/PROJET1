import pandas as pd
from xlwt import Workbook

lecture = pd.read_excel("projet_edacy.xlsx", index_col=0)

list_moyenne = lecture.Moyenne >= 10

print("La listes des eleves ayant la moyenne est la suivante:\n")

print(lecture[list_moyenne])
print("\n")

fichier1 = open("list_eleve_ayant_moyenne.xlsx","w")
r=lecture[list_moyenne]
fichier1.write(str(r))
fichier1.close()

list_plus_vingt = lecture.Age >= 20

print("La listes des eleves ayant plus de 20 ans est la suivante:\n ")

print(lecture[list_plus_vingt])
print("\n")

fichier2 = open("list_eleve_ayant_plus_20ans.xlsx", "w")
m = lecture[list_plus_vingt]
fichier2.write(str(m))
fichier2.close()

print("La moyenne de l'ecole est {} ".format(lecture['Moyenne'].mean(axis=0)) )

i = 0
j = 0

for v in lecture.Sexe:
    if v == 'f':
        i+=1
print("Le pourcentage de filles est {} %".format(100*float(i)/float(50)))
for v in lecture.Sexe:
    if v =='m':
        j+=1
print("Le pourcentage de garcons est {} %".format(100*float(j)/float(50)))

k=10

for l in lecture.Moyenne:
    if l>k:
        k=int(l)
        k=int(k)

region_forte= lecture.Moyenne == k

x: object
for x in lecture[region_forte].Region:
    print("La region ayant enregistrée la plus forte moyenne est {}".format(x))

book = Workbook()

feuil1 = book.add_sheet('feuille 1')

feuil1.write(0, 0, 'Moyenne')
feuil1.write(0, 1, 'Pourcentage filles')
feuil1.write(0, 2, 'Pourcentage garcons')
feuil1.write(0, 3, 'Région Forte')

ligne1 = feuil1.row(1)
ligne1.write(0, lecture['Moyenne'].mean(axis=0))
ligne1.write(1, 100*float(i)/float(50))
ligne1.write(2, 100*float(j)/float(50))
ligne1.write(3, x)

book.save('StatistiquesGlobales.xlsx')

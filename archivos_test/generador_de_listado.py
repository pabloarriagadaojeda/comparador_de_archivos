import pandas as pd
import random

first_names = ["Juan", "Maria", "Pedro", "Ana", "Luis", "Carmen", "Jose", "Laura", "Miguel", "Marta"]
last_names = ["Perez", "Gonzalez", "Rodriguez", "Fernandez", "Lopez", "Martinez", "Garcia", "Sanchez", "Diaz", "Hernandez"]

def generate_phone_number():
    return f"+56 9 {random.randint(10000000, 99999999)}"

people = []
for i in range(1, 201):
    first_name = random.choice(first_names)
    last_name = random.choice(last_names)
    name = f"{first_name} {last_name}"
    phone = generate_phone_number()
    people.append({"id": i, "nombre": name, "telefono": phone})

df = pd.DataFrame(people)

csv_file = "gente.csv"

df.to_csv(csv_file, index=False, encoding='utf-8')

print(f"El archivo {csv_file} ha sido creado exitosamente.")

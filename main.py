import pandas as pd

class Estudiante:
    def __init__(self, nombre, correo, cedula, semestre, materia, nota1, nota2):
        self.nombre = nombre
        self.correo = correo
        self.cedula = cedula
        self.semestre = semestre
        self.materia = materia
        self.nota1 = nota1
        self.nota2 = nota2

    def to_dict(self):
        return {
            "Nombre": self.nombre,
            "Correo": self.correo,
            "Cédula": self.cedula,
            "Semestre": self.semestre,
            "Materia": self.materia,
            "Nota 1": self.nota1,
            "Nota 2": self.nota2,
        }

class GestorEstudiantes:
    def __init__(self, ruta_csv):
        self.ruta_csv = ruta_csv
        self.estudiantes = []

    def cargar_datos_base(self):
        df = pd.read_csv(self.ruta_csv)
        df.columns = df.columns.str.strip()
        return df

    def recolectar_datos_completos(self, df_base):
        print("Introduce los datos adicionales para cada estudiante:\n")
        for _, fila in df_base.iterrows():
            print(f"Estudiante: {fila['Nombre']} ({fila['Correo']})")

            cedula = input("  Cédula: ") or None
            semestre = input("  Semestre: ") or None
            materia = input("  Materia: ") or None

            nota1_input = input("  Nota del Momento 1: ")
            nota1 = float(nota1_input) if nota1_input.strip() else None

            nota2_input = input("  Nota del Momento 2: ")
            nota2 = float(nota2_input) if nota2_input.strip() else None

            estudiante = Estudiante(
                nombre=fila['Nombre'],
                correo=fila['Correo'],
                cedula=cedula,
                semestre=semestre,
                materia=materia,
                nota1=nota1,
                nota2=nota2
            )

            self.estudiantes.append(estudiante)
            print("-" * 40)

    def guardar_en_excel(self, nombre_archivo="resources/notas_estudiantes.xlsx"):
        materias = {}
        for estudiante in self.estudiantes:
            if estudiante.materia not in materias:
                materias[estudiante.materia] = []
            materias[estudiante.materia].append(estudiante.to_dict())

        with pd.ExcelWriter(nombre_archivo) as writer:
            for materia, lista_estudiantes in materias.items():
                nombre_hoja = materia[:31] if materia else "SinMateria"
                df = pd.DataFrame(lista_estudiantes)
                df.to_excel(writer, sheet_name=nombre_hoja, index=False)

        print(f"\nArchivo Excel '{nombre_archivo}' creado exitosamente.")

if __name__ == "__main__":
    ruta_csv = "resources/estudiantes.csv"
    gestor = GestorEstudiantes(ruta_csv)
    df_base = gestor.cargar_datos_base()
    gestor.recolectar_datos_completos(df_base)
    gestor.guardar_en_excel()
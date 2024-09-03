import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
import os


def procesar_excel():
    
    Tk().withdraw()

    
    ruta_archivo = askopenfilename(
        title="Selecciona el archivo Excel - MALLA OPS",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")]
    )

   
    if ruta_archivo:
        
        hojas = pd.ExcelFile(ruta_archivo).sheet_names

        # Buscar la hoja que coincida con el mes sin importar M o m
        hoja_mes = next((hoja for hoja in hojas if hoja.lower() == 'agosto'), None)

        if hoja_mes:
            # Leer la hoja
            df_mes = pd.read_excel(ruta_archivo, sheet_name=hoja_mes)  #NOMBRE DE LA HOJA #modificar nombre M AND m

        
        columnas_deseadas = ['BOOKING SERVICIO', 'total_company_final_cost']
        df_filtrado = df_mes[columnas_deseadas]

        
        df_filtrado['total_company_final_cost'] = df_filtrado['total_company_final_cost'].astype(str).str.replace('.', ',', regex=False)

        ruta_csv = asksaveasfilename( #GUARDAR EN CARPETA DEL PROYECTO
            title="Guardar EN CARPETA",
            defaultextension=".csv",
            filetypes=[("Archivo CSV", "*.csv")]
        )

        if ruta_csv:
   
            df_filtrado.to_csv(ruta_csv, index=False)
            print(f"GUARDADO EN {ruta_csv}")
        else:
            print("ERROR")
    else:
        print("NO HA SELECCIONADO LA MALLA")

# FUNCION DE COMBINAR LOS ARCHIVOS
def combinar_csv():
    
    Tk().withdraw()

    #
    carpeta = askdirectory(title="Selecciona la carpeta con los archivos CSV")

    
    if carpeta:
        
        try:
            # LEE EL QUERY
            ruta_facturacion = os.path.join(carpeta, 'FACTURACION.csv')
            df_facturacion = pd.read_csv(ruta_facturacion)
            df_facturacion_filtrado = df_facturacion[['Nombre Compañia','Booking ID', 'GMV']] 

            # Leer LA MALLA
            ruta_malla = os.path.join(carpeta, 'MALLA.csv')
            df_malla = pd.read_csv(ruta_malla)

            # CRUCE
            df_combined = pd.merge(df_facturacion_filtrado, df_malla, left_on='Booking ID', right_on='BOOKING SERVICIO', how='left')
            

            ruta_csv_guardar = asksaveasfilename(
                title="Guardar archivo final",
                defaultextension=".csv",
                filetypes=[("Archivo CSV", "*.csv")]
            )

            
            if ruta_csv_guardar:
               
                df_combined.to_csv(ruta_csv_guardar, index=False)
                print(f"El ARCHIVO FINAL se ha guardado correctamente en {ruta_csv_guardar}")
            else:
                print("No selecciono ninguna ubicación para guardar el archivo CSV.")
        except FileNotFoundError as e:
            print(f"Error en archivos : {e}")
        except KeyError as e:
            print(f"Columnas no encontradas: {e}")
        except Exception as e:
            print(f"Error: {e}")
    else:
        print("No selecciono ninguna carpeta.")

#funcion final para ejecutar funciones
def main():
    procesar_excel()  # Procesar mala
    combinar_csv()    # Cruce

main()

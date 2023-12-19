import pandas as pd
import tkinter as tk
from tkinter import messagebox



def load_employee_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        print(f"File not found at path: {file_path}")
        return None


def extract_numeric_id(badge_number):
    # Assuming the numeric ID is always in the format "R000<number>A"
    return int(''.join(filter(str.isdigit, str(badge_number))))


def mark_gift_given(employee_data, badge_number, file_path):
    numeric_id = extract_numeric_id(badge_number)
    try:
        index = employee_data.index[employee_data['Numero'] == numeric_id].tolist()[
            0]
        if int(employee_data.at[index, 'Entrega_Regalo']) == 1:
            messagebox.showinfo(
                "Info", f"""
                    -----------------------------------------
                    -----------------------------------------
                    -----------------------------------------
                    El empleado con el ID {numeric_id} ya recibió el regalo.
                    -----------------------------------------
                    -----------------------------------------
                    -----------------------------------------
                    """)
        else:
            employee_data.at[index, 'Entrega_Regalo'] = 1
            employee_data.to_excel(file_path, index=False)
            messagebox.showinfo(
                "Success", f"Regalo marcado como entregado para el empleado con ID {numeric_id}")
    except IndexError:
        messagebox.showwarning(
            "Error", f"No se encontró empleado con el ID {numeric_id}")


class GiftApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Gift Tracker")

        self.label = tk.Label(root, text="Ingrese el número de empleado:")
        self.label.pack(pady=10)

        self.entry = tk.Entry(root, width=20)
        self.entry.pack(pady=10)

        self.button = tk.Button(
            root, text="Marcar Regalo Entregado", command=self.mark_gift)
        self.button.pack(pady=10)

        self.load_data_button = tk.Button(
            root, text="Cargar Datos de Empleados", command=self.load_data)
        self.load_data_button.pack(pady=30)

    def load_data(self):
        excel_file_path = 'employee_data.xlsx'
        self.employee_data = load_employee_data(excel_file_path)
        if self.employee_data is not None:
            messagebox.showinfo(
                "Success", "Datos de empleados cargados exitosamente")
            print("Employee data loaded successfully.")
        else:
            messagebox.showwarning(
                "Error", "No se encontró el archivo de excel en la misma carpeta que el programa.")
            print("Error: Employee data not loaded. File not found.")

    def mark_gift(self):
        if hasattr(self, 'employee_data'):
            badge_number = str(self.entry.get())  # Convert to string
            if badge_number.strip():  # Check if the input is not empty or contains only whitespace
                numeric_id = extract_numeric_id(badge_number)
                mark_gift_given(self.employee_data, numeric_id, 'employee_data.xlsx')
                self.entry.delete(0, tk.END)  # Clear the entry widget
            else:
                messagebox.showwarning(
                    "Error", "Ingrese un número de empleado antes de marcar el regalo.")
        else:
            messagebox.showwarning(
                "Error", "Datos de empleados no cargados. Haga clic en 'Cargar Datos de Empleados' primero.")


if __name__ == "__main__":
    root = tk.Tk()
    app = GiftApp(root)
    root.mainloop()

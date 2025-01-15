import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
import pandas as pd
import uuid
import os
from fpdf import FPDF
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class TimeTrackingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SUIVI DE TEMPS DE PRODUCTION")
        self.root.geometry("800x600")  # Set the window size

        self.start_time_var = tk.StringVar()
        self.end_time_var = tk.StringVar()
        self.downtime_var = tk.StringVar()
        self.machine_name_var = tk.StringVar()
        self.order_num_var = tk.StringVar()
        self.followup_date_var = tk.StringVar()  # Nouvelle variable pour la date de suivi
        self.filter_machine_var = tk.StringVar()
        self.machines = ["CTL1250", "PBM400", "REF2000", "PROFILEUSE C&Z", "PLIEUSE", "CTL1600", "REF1250", "PLASMA",
                         "PROFILEUSE Italie", "PBM200", "T-ONDULEE", "MICRO NERVUREE", "TOITESCO", "PROFILEUSE TURQUIE",
                         "HERMAK", "METAL DEPLOYE", "PERFORATRICE", "OSSATURE", "PROFILEUSE LAMES ET GLISSIERE",
                         "PANNE C N°2", "FLASQUE", "TOTAL MACHINES"]

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(pady=10, expand=True, fill=tk.BOTH)

        self.entry_frame = ttk.Frame(self.notebook)
        self.report_frame = ttk.Frame(self.notebook)
        self.graphical_report_frame = ttk.Frame(self.notebook)

        self.notebook.add(self.entry_frame, text="SAISIE DE TEMPS")
        self.notebook.add(self.report_frame, text="RAPPORT")
        self.notebook.add(self.graphical_report_frame, text="RAPPORT GRAPHIQUE")

        self.create_entry_tab()
        self.create_report_tab()
        self.create_graphical_report_tab()

        # Add the function to save data automatically when the application closes
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_entry_tab(self):
        entry_container = ttk.Frame(self.entry_frame, padding=10)
        entry_container.pack(fill=tk.BOTH, expand=True)

        tk.Label(entry_container, text="Machine :", bg="white").grid(row=0, column=0, pady=5, sticky=tk.W)
        self.machine_name_menu = tk.OptionMenu(entry_container, self.machine_name_var, *self.machines)
        self.machine_name_menu.grid(row=0, column=1, pady=5, sticky=tk.EW)

        tk.Label(entry_container, text="N° ordre de fabrication:", bg="white").grid(row=1, column=0, pady=5, sticky=tk.W)
        self.order_num_entry = tk.Entry(entry_container, textvariable=self.order_num_var)
        self.order_num_entry.grid(row=1, column=1, pady=5, sticky=tk.EW)

        tk.Label(entry_container, text="Temps de fonctionnement:", bg="white").grid(row=2, column=0, pady=5, sticky=tk.W)
        self.start_time_entry = tk.Entry(entry_container, textvariable=self.start_time_var)
        self.start_time_entry.grid(row=2, column=1, pady=5, sticky=tk.EW)

        tk.Label(entry_container, text="Temps de réglage:", bg="white").grid(row=3, column=0, pady=5, sticky=tk.W)
        self.end_time_entry = tk.Entry(entry_container, textvariable=self.end_time_var)
        self.end_time_entry.grid(row=3, column=1, pady=5, sticky=tk.EW)

        tk.Label(entry_container, text="Temps d'arrêts:", bg="white").grid(row=4, column=0, pady=5, sticky=tk.W)
        self.downtime_entry = tk.Entry(entry_container, textvariable=self.downtime_var)
        self.downtime_entry.grid(row=4, column=1, pady=5, sticky=tk.EW)

        tk.Label(entry_container, text="Date de Suivi:", bg="white").grid(row=5, column=0, pady=5, sticky=tk.W)
        self.followup_date_entry = DateEntry(entry_container, textvariable=self.followup_date_var, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.followup_date_entry.grid(row=5, column=1, pady=5, sticky=tk.EW)

        button_frame = ttk.Frame(entry_container)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)

        tk.Button(button_frame, text="Add Entry", command=self.add_entry, bg="#4285F4", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Save to Excel", command=self.save_to_excel, bg="#34A853", fg="white").pack(side=tk.LEFT, padx=10)

        self.log_text = tk.Text(entry_container, height=10, bg="white", fg="black")
        self.log_text.grid(row=7, column=0, columnspan=2, pady=10, sticky=tk.NSEW)

        self.entries = []

    def create_report_tab(self):
        report_container = ttk.Frame(self.report_frame, padding=10)
        report_container.pack(fill=tk.BOTH, expand=True)

        tk.Label(report_container, text="Start Date:", bg="white").grid(row=0, column=0, pady=5, sticky=tk.W)
        self.start_date_entry = DateEntry(report_container, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.start_date_entry.grid(row=0, column=1, pady=5, sticky=tk.EW)

        tk.Label(report_container, text="End Date:", bg="white").grid(row=1, column=0, pady=5, sticky=tk.W)
        self.end_date_entry = DateEntry(report_container, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.end_date_entry.grid(row=1, column=1, pady=5, sticky=tk.EW)

        tk.Label(report_container, text="Filter by Machine:", bg="white").grid(row=2, column=0, pady=5, sticky=tk.W)
        self.filter_machine_menu = tk.OptionMenu(report_container, self.filter_machine_var, *self.machines)
        self.filter_machine_menu.grid(row=2, column=1, pady=5, sticky=tk.EW)

        button_frame = ttk.Frame(report_container)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        tk.Button(button_frame, text="Generate Report", command=self.generate_report, bg="#4285F4", fg="white").pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Export to PDF", command=self.export_to_pdf, bg="#FF5733", fg="white").pack(side=tk.LEFT, padx=10)

        self.report_text = tk.Text(report_container, height=10, bg="white", fg="black")
        self.report_text.grid(row=4, column=0, columnspan=2, pady=10, sticky=tk.NSEW)

    def create_graphical_report_tab(self):
        graphical_report_container = ttk.Frame(self.graphical_report_frame, padding=10)
        graphical_report_container.pack(fill=tk.BOTH, expand=True)

        tk.Label(graphical_report_container, text="Start Date:", bg="white").grid(row=0, column=0, pady=5, sticky=tk.W)
        self.graphical_start_date_entry = DateEntry(graphical_report_container, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.graphical_start_date_entry.grid(row=0, column=1, pady=5, sticky=tk.EW)

        tk.Label(graphical_report_container, text="End Date:", bg="white").grid(row=1, column=0, pady=5, sticky=tk.W)
        self.graphical_end_date_entry = DateEntry(graphical_report_container, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.graphical_end_date_entry.grid(row=1, column=1, pady=5, sticky=tk.EW)

        tk.Label(graphical_report_container, text="Filter by Machine:", bg="white").grid(row=2, column=0, pady=5, sticky=tk.W)
        self.graphical_filter_machine_menu = tk.OptionMenu(graphical_report_container, self.filter_machine_var, *self.machines)
        self.graphical_filter_machine_menu.grid(row=2, column=1, pady=5, sticky=tk.EW)

        button_frame = ttk.Frame(graphical_report_container)
        button_frame.grid(row=3, column=0, columnspan=2, pady=20)

        tk.Button(button_frame, text="Generate Graphical Report", command=self.generate_graphical_report, bg="#4285F4", fg="white").pack(side=tk.LEFT, padx=10)

        self.graphical_report_canvas = None

    def add_entry(self):
        # Vérifier si une machine est sélectionnée
        machine_name = self.machine_name_var.get()
        order_num = self.order_num_var.get()

        # Highlight required fields if empty
        if not machine_name:
            self.machine_name_menu.config(bg="red")
        else:
            self.machine_name_menu.config(bg="white")

        if not order_num:
            self.order_num_entry.config(bg="red")
        else:
            self.order_num_entry.config(bg="white")

        if not machine_name or not order_num:
            messagebox.showerror("Error", "Please fill in all required fields.")
            return

        # Récupérer les autres données
        start_time = self.start_time_var.get()
        end_time = self.end_time_var.get()
        downtime = self.downtime_var.get()
        declaration_num = str(uuid.uuid4())
        followup_date = self.followup_date_var.get()  # Date de suivi

        # Créer l'entrée avec la nouvelle date de suivi
        entry = {
            "Machine Name": machine_name,
            "N° de déclaration": declaration_num,
            "N° ordre de fabrication": order_num,
            "Temps de fonctionnement": start_time,
            "Temps de réglage": end_time,
            "Temps d'arrêts": downtime,
            "Date": pd.Timestamp.now().strftime("%Y-%m-%d"),
            "Date de suivi": followup_date  # Ajouter la date de suivi
        }
        self.entries.append(entry)
        self.log_text.insert(tk.END, f"Added entry: {entry}\n")
        self.log_text.see(tk.END)

        # Reset fields after adding entry
        self.machine_name_var.set("")
        self.order_num_var.set("")
        self.start_time_var.set("")
        self.end_time_var.set("")
        self.downtime_var.set("")
        self.followup_date_var.set("")

        # Display message with declaration number
        messagebox.showinfo("Success", f"Entry added successfully with Declaration Number: {declaration_num}")

    def save_to_excel(self):
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        file_path = os.path.join(desktop_path, "time_tracking_entries.xlsx")

        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            new_df = pd.DataFrame(self.entries)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        else:
            combined_df = pd.DataFrame(self.entries)

        combined_df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", f"Data saved to {file_path}")

    def on_close(self):
        self.save_to_excel()
        self.root.quit()

    def generate_report(self):
        start_date = pd.to_datetime(self.start_date_entry.get_date())
        end_date = pd.to_datetime(self.end_date_entry.get_date())
        filter_machine = self.filter_machine_var.get()

        print(f"Start Date: {start_date}")
        print(f"End Date: {end_date}")

        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        file_path = os.path.join(desktop_path, "time_tracking_entries.xlsx")

        if not os.path.exists(file_path):
            messagebox.showerror("Error", "No data file found.")
            return

        df = pd.read_excel(file_path)

        # Convert the 'Date' column to datetime
        df['Date'] = pd.to_datetime(df['Date'])

        print("DataFrame contents:")
        print(df)

        filtered_entries = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        if filter_machine == "TOTAL MACHINES":
            total_fonctionnement = filtered_entries["Temps de fonctionnement"].astype(float).sum()
            total_reglage = filtered_entries["Temps de réglage"].astype(float).sum()
            total_arrets = filtered_entries["Temps d'arrêts"].astype(float).sum()

            report = (
                f"Total Temps de fonctionnement: {total_fonctionnement} minutes\n"
                f"Total Temps de réglage: {total_reglage} minutes\n"
                f"Total Temps d'arrêts: {total_arrets} minutes\n"
            )
        else:
            if filter_machine:
                filtered_entries = filtered_entries[filtered_entries['Machine Name'] == filter_machine]

            print("Filtered Entries:")
            print(filtered_entries)

            if filtered_entries.empty:
                messagebox.showerror("Error", "No valid entries to generate report.")
                return

            report = f"Report from {start_date.date()} to {end_date.date()}\n\n"

            grouped = filtered_entries.groupby('Machine Name')

            for machine, group in grouped:
                total_fonctionnement = 0
                total_reglage = 0
                total_arrets = 0

                for _, entry in group.iterrows():
                    fonctionnement = int(entry.get("Temps de fonctionnement", 0))
                    reglage = int(entry.get("Temps de réglage", 0))
                    arrets = int(entry.get("Temps d'arrêts", 0))

                    print(f"Machine: {machine}, Fonctionnement: {fonctionnement}, Reglage: {reglage}, Arrets: {arrets}")

                    total_fonctionnement += fonctionnement
                    total_reglage += reglage
                    total_arrets += arrets

                total_time = total_fonctionnement + total_reglage + total_arrets

                if total_time > 0:
                    fonctionnement_percentage = (total_fonctionnement / total_time) * 100
                    reglage_percentage = (total_reglage / total_time) * 100
                    arrets_percentage = (total_arrets / total_time) * 100

                    report += (
                        f"Machine: {machine}\n"
                        f"Total Temps de fonctionnement: {total_fonctionnement} minutes ({fonctionnement_percentage:.2f}%)\n"
                        f"Total Temps de réglage: {total_reglage} minutes ({reglage_percentage:.2f}%)\n"
                        f"Total Temps d'arrêts: {total_arrets} minutes ({arrets_percentage:.2f}%)\n\n"
                    )

        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)

    def export_to_pdf(self):
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        file_path = os.path.join(desktop_path, "time_tracking_report.pdf")

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)

        report_text = self.report_text.get("1.0", tk.END)
        for line in report_text.split('\n'):
            pdf.cell(200, 10, txt=line, ln=True)

        pdf.output(file_path)
        messagebox.showinfo("Success", f"Report exported to {file_path}")

    def generate_graphical_report(self):
        start_date = pd.to_datetime(self.graphical_start_date_entry.get_date())
        end_date = pd.to_datetime(self.graphical_end_date_entry.get_date())
        filter_machine = self.filter_machine_var.get()

        print(f"Start Date: {start_date}")
        print(f"End Date: {end_date}")

        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        file_path = os.path.join(desktop_path, "time_tracking_entries.xlsx")

        if not os.path.exists(file_path):
            messagebox.showerror("Error", "No data file found.")
            return

        df = pd.read_excel(file_path)

        # Convert the 'Date' column to datetime
        df['Date'] = pd.to_datetime(df['Date'])

        print("DataFrame contents:")
        print(df)

        filtered_entries = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

        if filter_machine == "TOTAL MACHINES":
            total_fonctionnement = filtered_entries["Temps de fonctionnement"].astype(float).sum()
            total_reglage = filtered_entries["Temps de réglage"].astype(float).sum()
            total_arrets = filtered_entries["Temps d'arrêts"].astype(float).sum()

            machine_names = ["TOTAL MACHINES"]
            fonctionnement_times = [total_fonctionnement]
            reglage_times = [total_reglage]
            arrets_times = [total_arrets]
        else:
            if filter_machine:
                filtered_entries = filtered_entries[filtered_entries['Machine Name'] == filter_machine]

            print("Filtered Entries:")
            print(filtered_entries)

            if filtered_entries.empty:
                messagebox.showerror("Error", "No valid entries to generate report.")
                return

            grouped = filtered_entries.groupby('Machine Name')

            machine_names = []
            fonctionnement_times = []
            reglage_times = []
            arrets_times = []

            for machine, group in grouped:
                total_fonctionnement = 0
                total_reglage = 0
                total_arrets = 0

                for _, entry in group.iterrows():
                    fonctionnement = int(entry.get("Temps de fonctionnement", 0))
                    reglage = int(entry.get("Temps de réglage", 0))
                    arrets = int(entry.get("Temps d'arrêts", 0))

                    total_fonctionnement += fonctionnement
                    total_reglage += reglage
                    total_arrets += arrets

                machine_names.append(machine)
                fonctionnement_times.append(total_fonctionnement)
                reglage_times.append(total_reglage)
                arrets_times.append(total_arrets)

        fig, ax = plt.subplots()
        bar_width = 0.25
        index = range(len(machine_names))

        bar1 = plt.bar(index, fonctionnement_times, bar_width, label='Temps de fonctionnement')
        bar2 = plt.bar([i + bar_width for i in index], reglage_times, bar_width, label='Temps de réglage')
        bar3 = plt.bar([i + 2 * bar_width for i in index], arrets_times, bar_width, label='Temps d\'arrêts')

        plt.xlabel('Machines')
        plt.ylabel('Time (minutes)')
        plt.title('Machine Time Report')
        plt.xticks([i + bar_width for i in index], machine_names, rotation=45)
        plt.legend()

        if self.graphical_report_canvas:
            self.graphical_report_canvas.get_tk_widget().destroy()

        self.graphical_report_canvas = FigureCanvasTkAgg(fig, master=self.graphical_report_frame)
        self.graphical_report_canvas.draw()
        self.graphical_report_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeTrackingApp(root)
    root.mainloop()

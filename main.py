import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import matplotlib.pyplot as plt

class DSSApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Decision Support System")
        self.root.state("zoomed")
        self.root.configure(bg="pink")

        self.file_path = tk.StringVar()
        self.df = None
        self.columns = [
            'Project Description', 'Project Type', 'Project Manager',
            'Region', 'Department', 'Project Cost', 'Project Benefit',
            'Complexity', 'Status', 'Completion%', 'Project duration',
            'Start Date', 'End Date', 'Business criticality'
        ]

        self.criteria_options = {
            "Project Cost": [
                ">=5mln",
                "from 4.9mln to 3mln",
                "from 2.9mln to 1mln",
                "<1mln"
            ],
            "Project Benefit": [
                ">=10mln",
                "from 9.9mln to 5mln",
                "from 4.9mln to 1mln",
                "<1mln"
            ],
            "Complexity": [
                "High",
                "Medium",
                "Low"
            ],
            "Completion%": [
                ">=75%",
                "from 75 to 50%",
                "from 50 to 25%",
                "<25%"
            ],
            "Project duration": [
                ">= 9 months",
                ">= 6 months",
                ">= 2 months",
                "< 2 months"
            ],
            "End Date": [
                "today<=60",
                "today from 60 to 90",
                "today from 90 to 180",
                "today>180"
            ],
            "Business criticality": [
                "Legally required/Key market project",
                "Function/Business priority (OMP)",
                "Globaly required",
                "Acceptable workaround available – “nice to have” project"
            ]
        }

        self.score_options = {
            "4, 3, 2, 1": [4, 3, 2, 1],
            "0, 0, 0, 0": [0, 0, 0, 0]
        }
        self.user_selections = {}

        self.create_upload_view()

    def create_upload_view(self):
        self.root.geometry("600x300")
        tk.Label(self.root, text="Select a CSV file to upload:", font=('Helvetica', 14)).pack(pady=10)
        tk.Entry(self.root, textvariable=self.file_path, width=70, state='readonly').pack(pady=5)
        tk.Button(self.root, text="Browse", command=self.browse_file).pack(pady=5)
        tk.Button(self.root, text="Upload", command=self.upload_file).pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if file_path:
            self.file_path.set(file_path)

    def upload_file(self):
        path = self.file_path.get()
        if not path:
            messagebox.showerror("Error", "Please select a file.")
            return

        try:
            self.df = pd.read_csv(path)

            missing_columns = [col for col in self.columns if col not in self.df.columns]
            if missing_columns:
                messagebox.showerror("Error", f"File is missing required columns: {', '.join(missing_columns)}")
                return

            print(self.df.head())

            messagebox.showinfo("Success", "CSV file uploaded and read successfully.")

            self.create_score_selection_view()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def create_score_selection_view(self):
        self.root.geometry("800x600")

        canvas = tk.Canvas(self.root, borderwidth=0, background="#f0f0f0")
        scroll_y = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scroll_y.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        frame = tk.Frame(canvas, background="#f0f0f0")
        canvas.create_window((0, 0), window=frame, anchor="nw")

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)

        for criterion, options in self.criteria_options.items():
            criterion_label = tk.Label(frame, text=criterion, font=('Helvetica', 12, 'bold'))
            criterion_label.pack(anchor='w', padx=10, pady=(10, 5))

            option_frame = tk.Frame(frame, padx=20, pady=10)
            option_frame.pack(anchor='w', padx=10, pady=5)

            values_label = tk.Label(option_frame, text="Criterion Values:")
            values_label.grid(row=0, column=0, sticky='w', padx=10)

            for idx, option in enumerate(options):
                value_label = tk.Label(option_frame, text=option)
                value_label.grid(row=1+idx, column=0, sticky='w', padx=10, pady=2)

            scoring_label = tk.Label(option_frame, text="Select scoring option:")
            scoring_label.grid(row=0, column=1, sticky='w', padx=10)

            for row, score in enumerate(self.score_options["4, 3, 2, 1"], start=1):
                score_label = tk.Label(option_frame, text=str(score))
                score_label.grid(row=row, column=1, sticky='w', padx=10, pady=2)

            for row, score in enumerate(self.score_options["0, 0, 0, 0"], start=1):
                score_label = tk.Label(option_frame, text=str(score))
                score_label.grid(row=row, column=2, sticky='w', padx=10, pady=2)

            selected_option = tk.StringVar(value="4, 3, 2, 1")
            for row, (option_text, _) in enumerate(self.score_options.items(), start=1):
                radio_button = tk.Radiobutton(option_frame, text=option_text, variable=selected_option, value=option_text)
                radio_button.grid(row=row, column=3, sticky='w', padx=10, pady=2)

            self.user_selections[criterion] = selected_option

        submit_button = tk.Button(frame, text="Submit", command=self.submit_form)
        submit_button.pack(pady=20)

        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def submit_form(self):
        result = {}
        criteria_to_score = [
            'Project Cost', 'Project Benefit', 'Complexity',
            'Completion%', 'Project duration', 'End Date', 'Business criticality'
        ]

        for criterion in criteria_to_score:
            selected_option = self.user_selections[criterion].get()
            if selected_option == "4, 3, 2, 1":
                scores = [4, 3, 2, 1]
            elif selected_option == "0, 0, 0, 0":
                scores = [0, 0, 0, 0]
            else:
                return

            selected_score = {}
            for idx, option in enumerate(self.criteria_options[criterion]):
                selected_score[option] = scores[idx]

            result[criterion] = selected_score

            score_field_name = f"{criterion.lower().replace(' ', '_')}_score"
            self.df[score_field_name] = 0

            def score_project_cost(cost):
                try:
                    cost = float(str(cost).replace(',', ''))
                except ValueError:
                    cost = 0

                if cost >= 5000000:
                    return selected_score[">=5mln"]
                elif 3000000 <= cost < 5000000:
                    return selected_score["from 4.9mln to 3mln"]
                elif 1000000 <= cost < 3000000:
                    return selected_score["from 2.9mln to 1mln"]
                else:
                    return selected_score["<1mln"]

            def score_complexity(complexity):
                return selected_score.get(complexity, 0)

            def score_project_benefit(benefit):
                try:
                    benefit = float(str(benefit).replace(',', ''))
                except ValueError:
                    benefit = 0

                if benefit >= 10000000:
                    return selected_score[">=10mln"]
                elif 5000000 <= benefit < 10000000:
                    return selected_score["from 9.9mln to 5mln"]
                elif 1000000 <= benefit < 5000000:
                    return selected_score["from 4.9mln to 1mln"]
                else:
                    return selected_score["<1mln"]

            def score_completion(completion):
                try:
                    completion = float(str(completion).replace('%', ''))
                except ValueError:
                    completion = 0

                if completion >= 75:
                    return selected_score[">=75%"]
                elif 50 <= completion < 75:
                    return selected_score["from 75 to 50%"]
                elif 25 <= completion < 50:
                    return selected_score["from 50 to 25%"]
                else:
                    return selected_score["<25%"]

            def score_project_duration(duration):
                try:
                    duration = float(duration)
                except ValueError:
                    duration = 0

                if duration >= 9:
                    return selected_score[">= 9 months"]
                elif 6 <= duration < 9:
                    return selected_score[">= 6 months"]
                elif 2 <= duration < 6:
                    return selected_score[">= 2 months"]
                else:
                    return selected_score["< 2 months"]

            def score_end_date(end_date):
                try:
                    end_date = pd.to_datetime(end_date)
                except ValueError:
                    return 0

                today = datetime.today()
                days_to_end = (end_date - today).days
                if days_to_end <= 60:
                    return selected_score["today<=60"]
                elif 60 < days_to_end <= 90:
                    return selected_score["today from 60 to 90"]
                elif 90 < days_to_end <= 180:
                    return selected_score["today from 90 to 180"]
                else:
                    return selected_score["today>180"]

            def score_business_criticality(criticality):
                return selected_score.get(criticality, 0)

            for index, row in self.df.iterrows():
                if criterion == 'Project Cost':
                    project_value = row[criterion]
                    project_score = score_project_cost(project_value)
                elif criterion == 'Complexity':
                    complexity_value = row[criterion]
                    project_score = score_complexity(complexity_value)
                elif criterion == 'Project Benefit':
                    benefit_value = row[criterion]
                    project_score = score_project_benefit(benefit_value)
                elif criterion == 'Completion%':
                    completion_value = row[criterion]
                    project_score = score_completion(completion_value)
                elif criterion == 'Project duration':
                    duration_value = row[criterion]
                    project_score = score_project_duration(duration_value)
                elif criterion == 'End Date':
                    end_date_value = row[criterion]
                    project_score = score_end_date(end_date_value)
                elif criterion == 'Business criticality':
                    criticality_value = row[criterion]
                    project_score = score_business_criticality(criticality_value)
                else:
                    continue

                self.df.at[index, score_field_name] = project_score

        self.df['total_score'] = (
                self.df['project_cost_score'] +
                self.df['project_benefit_score'] +
                self.df['complexity_score'] +
                self.df['completion%_score'] +
                self.df['project_duration_score'] +
                self.df['end_date_score'] +
                self.df['business_criticality_score']
        )

        print(self.df.head())

        messagebox.showinfo("Form Submitted", "Scores have been successfully calculated and filled!")

        download_btn = tk.Button(self.root, text="Download CSV", command=self.download_csv)
        download_btn.pack(pady=10)

        show_top_10_btn = tk.Button(self.root, text="Show Top 10 Projects", command=self.show_top_10_projects)
        show_top_10_btn.pack(pady=10)

        button = tk.Button(self.root, text="Show Complexity Distribution", command=self.show_complexity_distribution)
        button.pack(pady=10)

    def show_complexity_distribution(self):
        grouped_data = self.df.groupby(['Complexity', 'Project Type']).size().unstack()

        plt.figure(figsize=(10, 6))
        grouped_data.plot(kind='bar', stacked=True, color=['red', 'lightgreen', 'lightcoral', 'lightskyblue'])
        plt.xlabel('Complexity')
        plt.ylabel('Number of Projects')
        plt.title('Distribution of Projects by Complexity and Project Type')
        plt.xticks(rotation=45)
        plt.legend(title='Project Type', loc='upper right')
        plt.tight_layout()

        plt.show()

    def download_csv(self):
        save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                 filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if save_path:
            self.df.to_csv(save_path, index=False)
            messagebox.showinfo("Success", f"CSV file has been saved to {save_path}")

    def show_top_10_projects(self):
        top_10_projects = self.df.nlargest(10, 'total_score')

        def calculate_developers(row):
            complexity = row['Complexity']
            project_type = row['Project Type']

            if complexity == 'High':
                if project_type == 'Web App':
                    return {'Web Developers': 4, 'Mobile Developers': 0, 'ML Developers': 0}
                elif project_type == 'Mobile App':
                    return {'Web Developers': 0, 'Mobile Developers': 4, 'ML Developers': 0}
                elif project_type == 'ML and AI App':
                    return {'Web Developers': 0, 'Mobile Developers': 0, 'ML Developers': 4}
                elif project_type == 'Web and Mobile':
                    return {'Web Developers': 4, 'Mobile Developers': 4, 'ML Developers': 0}
            elif complexity == 'Medium':
                if project_type == 'Web App':
                    return {'Web Developers': 2, 'Mobile Developers': 0, 'ML Developers': 0}
                elif project_type == 'Mobile App':
                    return {'Web Developers': 0, 'Mobile Developers': 2, 'ML Developers': 0}
                elif project_type == 'ML and AI App':
                    return {'Web Developers': 0, 'Mobile Developers': 0, 'ML Developers': 2}
                elif project_type == 'Web and Mobile':
                    return {'Web Developers': 2, 'Mobile Developers': 2, 'ML Developers': 0}
            elif complexity == 'Low':
                if project_type == 'Web App':
                    return {'Web Developers': 1, 'Mobile Developers': 0, 'ML Developers': 0}
                elif project_type == 'Mobile App':
                    return {'Web Developers': 0, 'Mobile Developers': 1, 'ML Developers': 0}
                elif project_type == 'ML and AI App':
                    return {'Web Developers': 0, 'Mobile Developers': 0, 'ML Developers': 1}
                elif project_type == 'Web and Mobile':
                    return {'Web Developers': 1, 'Mobile Developers': 1, 'ML Developers': 0}

            return {'Web Developers': 0, 'Mobile Developers': 0, 'ML Developers': 0}

        top_10_projects['Developers Needed'] = top_10_projects.apply(calculate_developers, axis=1)

        new_window = tk.Toplevel(self.root)
        new_window.title("Top 10 Priority Projects")
        new_window.state("zoomed")

        tk.Label(new_window, text="Top 10 Priority Projects", font=('Helvetica', 16, 'bold')).pack(pady=10)

        frame = tk.Frame(new_window)
        frame.pack(fill='both', expand=True)

        canvas = tk.Canvas(frame)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        previous_end_date = None

        for i, (index, row) in enumerate(top_10_projects.iterrows()):
            devs_needed = row['Developers Needed']
            text = (f"{i + 1}. Project Name: {row['Project Name']}, "
                    f"Project Manager: {row['Project Manager']}, "
                    f"Start: {row['Start Date']}, End: {row['End Date']}, "
                    f"Score: {row['total_score']}, "
                    f"Web Developers: {devs_needed['Web Developers']}, "
                    f"Mobile Developers: {devs_needed['Mobile Developers']}, "
                    f"ML Developers: {devs_needed['ML Developers']}")

            if previous_end_date and pd.to_datetime(row['End Date']) > pd.to_datetime(previous_end_date):
                previous_devs_needed = top_10_projects.iloc[i - 1]['Developers Needed']
                devs_needed['Web Developers'] += previous_devs_needed['Web Developers']
                devs_needed['Mobile Developers'] += previous_devs_needed['Mobile Developers']
                devs_needed['ML Developers'] += previous_devs_needed['ML Developers']

            tk.Label(scrollable_frame, text=text, font=('Helvetica', 12)).pack(anchor='w', padx=10, pady=5)

            previous_end_date = row['End Date']

if __name__ == "__main__":
    root = tk.Tk()
    app = DSSApp(root)
    root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox

def fetch_openalex_data(doi):
    import requests
    url = f"https://api.openalex.org/works/doi:{doi}"
    response = requests.get(url)
    if response.status_code == 200:
        data = response.json()
        title = data.get('display_name', 'N/A')
        authors = ', '.join([author['author']['display_name'] for author in data.get('authorships', [])])
        abstract_inverted_index = data.get('abstract_inverted_index', {})
        
        if abstract_inverted_index:
            max_index = max(max(indices) for indices in abstract_inverted_index.values())
            abstract_list = [''] * (max_index + 1)
            for word, indices in abstract_inverted_index.items():
                for index in indices:
                    abstract_list[index] = word
            abstract = ' '.join(abstract_list)
        else:
            abstract = 'N/A'
            
        return [doi, title, authors, abstract]
    else:
        return [doi, 'Error', 'Error', 'Error']

def process_file(input_file, output_format):
    import pandas as pd
    try:
        if input_file.endswith('.csv'):
            df = pd.read_csv(input_file)
        elif input_file.endswith('.xlsx'):
            df = pd.read_excel(input_file)
        else:
            raise ValueError("Unsupported file format")
        
        if 'DOI' not in df.columns:
            raise ValueError("The input file must contain a 'DOI' column.")
        
        def normalize_doi(doi):
            doi = doi.strip()
            if doi.startswith("https://www.doi.org/"):
                doi = doi[len("https://www.doi.org/"):]
            return doi

        df['DOI'] = df['DOI'].apply(normalize_doi)

        results = []
        for doi in df['DOI']:
            result = fetch_openalex_data(doi)
            results.append(result)
        
        output_df = pd.DataFrame(results, columns=['DOI', 'Title', 'Authors', 'Abstract'])
        
        output_file = filedialog.asksaveasfilename(defaultextension=f'.{output_format}',
                                                   initialfile=f'output.{output_format}',
                                                   filetypes=[(f'{output_format.upper()} files', f'*.{output_format}')])
        if output_file:
            if output_format == 'csv':
                output_df.to_csv(output_file, index=False)
            elif output_format == 'xlsx':
                output_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"File saved successfully as {output_file}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def select_and_process_file():
    input_file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")])
    if input_file:
        output_format = output_format_var.get()
        process_file(input_file, output_format)

root = tk.Tk()
root.title("Abstract App")

tk.Label(root, text="Select Output Format:").pack(pady=10)
output_format_var = tk.StringVar(value="csv")
tk.Radiobutton(root, text="Comma-separated", variable=output_format_var, value="csv").pack()
tk.Radiobutton(root, text="Excel", variable=output_format_var, value="xlsx").pack()

tk.Button(root, text="Select Input File and Process", command=select_and_process_file).pack(pady=20)

if __name__ == "__main__":
    root.mainloop()

import random
import openpyxl
import os

class ExcelPicker:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": os.path.join("pony_char_list.xlsx")}),
                "sheet_name": ("STRING", {"default": "Female character list"}),
                "row_number": ("INT", {"default": 3, "min": 1, "max": 10000}),
                "num_outputs": ("INT", {"default": 5, "min": 1, "max": 20}),
                "prefix": ("STRING", {"default": "score_9, score_8_up, score_7_up"}),
                "seed_mode": (["randomize", "fixed"],),
                "seed": ("INT", {"default": 0, "min": 0, "max": 0xffffffffffffffff}),
            },
        }

    # Always return 20 string outputs and one integer seed
    RETURN_TYPES = tuple("STRING" for _ in range(20)) + ("INT",)
    RETURN_NAMES = tuple(f"output_{i+1}" for i in range(20)) + ("seed",)
    FUNCTION = "pick_prompt"
    CATEGORY = "DanOrAI-Excel"

    def pick_prompt(self, excel_path, sheet_name, row_number, num_outputs, prefix, seed_mode, seed):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_excel_path = os.path.join(current_dir, excel_path)
        
        if not os.path.exists(full_excel_path):
            raise FileNotFoundError(f"Excel file not found: {full_excel_path}")

        workbook = openpyxl.load_workbook(full_excel_path, read_only=True)
        
        if sheet_name not in workbook.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")

        sheet = workbook[sheet_name]
        
        if row_number > sheet.max_row:
            raise ValueError(f"Row {row_number} exceeds the number of rows in the sheet")

        # Convert cell values to strings and filter out empty ones
        row_values = [str(cell.value) for cell in sheet[row_number] if cell.value is not None]
        workbook.close()

        if not row_values:
            raise ValueError(f"No valid values found in row {row_number}")
        
        # Set the seed for random selection if needed
        if seed_mode == "randomize":
            seed = random.randint(0, 0xffffffffffffffff)
        random.seed(seed)

        # Select up to num_outputs values, pad with empty strings if not enough values
        selected_values = row_values[:num_outputs] + ["" for _ in range(max(0, num_outputs - len(row_values)))]
        # Ensure the list is exactly 20 items long
        if len(selected_values) < 20:
            selected_values += [""] * (20 - len(selected_values))
        
        # Process the prefix and format each output
        prefix_list = [p.strip() for p in prefix.split(',')]
        formatted_outputs = [
            ", ".join(prefix_list) + (", " + value if value else "")
            for value in selected_values
        ]

        return tuple(formatted_outputs) + (seed,)

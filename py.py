import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Укладка мультипалет")
        self.geometry("900x600")

        # Данные
        self.ozon_data = {}        # Артикул -> кол-во
        self.dimensions_data = {}  # Артикул -> размеры
        self.all_base = {}         # Артикул -> всё вместе

        # Интерфейс
        ttk = tk.ttk if hasattr(tk, 'ttk') else tk
        tk.Button(self, text="Загрузить накладную с OZON", command=self.load_ozon_file).pack(pady=5)
        tk.Button(self, text="Загрузить файл с размерами товаров", command=self.load_dimensions_file).pack(pady=5)

        frame = tk.Frame(self)
        frame.pack(pady=5)
        tk.Label(frame, text="Макс. высота палеты (см):").pack(side=tk.LEFT)
        self.height_entry = tk.Entry(frame)
        self.height_entry.insert(0, "180")
        self.height_entry.pack(side=tk.LEFT)

        tk.Button(self, text="Рассчитать палеты", command=self.calculate_pallets).pack(pady=10)

        self.result_text = tk.Text(self, height=25)
        self.result_text.pack(fill=tk.BOTH, expand=True)

    def load_ozon_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2):
            article = str(row[3].value).strip()  # 4-й столбец
            quantity = row[5].value              # 6-й столбец
            if article and quantity:
                self.ozon_data[article] = int(quantity)
        messagebox.showinfo("Успех", "Файл накладной загружен.")

    def load_dimensions_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path:
            return
        wb = openpyxl.load_workbook(path)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2):
            if any(cell.value is None for cell in row[:5]):
                continue
            article = str(row[0].value).strip()
            try:
                width = float(row[1].value)
                length = float(row[2].value)
                height = float(row[3].value)
                weight = float(row[4].value)
            except ValueError:
                continue
            self.dimensions_data[article] = {
                'width': width,
                'length': length,
                'height': height,
                'weight': weight
            }
        messagebox.showinfo("Успех", "Файл с размерами загружен.")
        self.merge_data()

    def merge_data(self):
        self.all_base.clear()
        for article, qty in self.ozon_data.items():
            if article in self.dimensions_data:
                self.all_base[article] = {
                    **self.dimensions_data[article],
                    'quantity': qty
                }

        print(self.all_base)
        
    def calculate_pallets(self):
        try:
            max_height = float(self.height_entry.get() or 180)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректную высоту палеты.")
            return
    
        pallet_width = 80  # см
        pallet_length = 120  # см
    
        from copy import deepcopy
    
        items = deepcopy(self.all_base)
        for item in items.values():
            item['remaining'] = item['quantity']
    
        def place_layer(layer_items):
            layer_plan = []
            grid = [[False] * int(pallet_length) for _ in range(int(pallet_width))]
    
            def fits(x, y, w, l):
                if x + w > pallet_width or y + l > pallet_length:
                    return False
                for i in range(int(x), int(x + w)):
                    for j in range(int(y), int(y + l)):
                        if grid[i][j]:
                            return False
                return True
    
            def occupy(x, y, w, l):
                for i in range(int(x), int(x + w)):
                    for j in range(int(y), int(y + l)):
                        grid[i][j] = True
    
            for article, data in layer_items.items():
                w1, l1 = data['width'], data['length']
                w2, l2 = l1, w1
                count = data['remaining']
                placed = 0
                for rotation in [(w1, l1), (w2, l2)]:
                    w, l = rotation
                    for x in range(int(pallet_width - w) + 1):
                        for y in range(int(pallet_length - l) + 1):
                            if fits(x, y, w, l) and count > 0:
                                occupy(x, y, w, l)
                                layer_plan.append({
                                    'article': article,
                                    'x': x,
                                    'y': y,
                                    'width': w,
                                    'length': l
                                })
                                placed += 1
                                count -= 1
                layer_items[article]['remaining'] -= placed
            return layer_plan
    
        pallets = []
        pallet_number = 1
    
        while any(i['remaining'] > 0 for i in items.values()):
            current_pallet = []
            current_height = 0
            layer_num = 1
            total_weight = 0
    
            while True:
                layer_items = {
                    k: v for k, v in items.items() if v['remaining'] > 0
                }
                if not layer_items:
                    break
                
                layer_plan = place_layer(layer_items)
    
                if not layer_plan:
                    break  # ничего не удалось уложить
                
                max_layer_height = 0
                layer_summary = {}
    
                for box in layer_plan:
                    art = box['article']
                    h = items[art]['height']
                    w = items[art]['weight']
                    max_layer_height = max(max_layer_height, h)
                    if art not in layer_summary:
                        layer_summary[art] = {'count': 0, 'weight': 0}
                    layer_summary[art]['count'] += 1
                    layer_summary[art]['weight'] += w
    
                if current_height + max_layer_height > max_height:
                    # Откат
                    for box in layer_plan:
                        items[box['article']]['remaining'] += 1
                    break
                
                for art, d in layer_summary.items():
                    current_pallet.append({
                        'article': art,
                        'count': d['count'],
                        'layer': layer_num
                    })
                    total_weight += d['weight']
    
                current_height += max_layer_height
                layer_num += 1
    
            pallets.append({
                'pallet_number': pallet_number,
                'layers': current_pallet,
                'weight': round(total_weight, 2)
            })
            pallet_number += 1
    
        # Вывод
        self.result_text.delete(1.0, tk.END)
        output = []
        for pallet in pallets:
            self.result_text.insert(tk.END, f"Палета {pallet['pallet_number']}:\n")
            output.append(f"Палета {pallet['pallet_number']}:")
            for layer in pallet['layers']:
                line = f"  - Артикул: {layer['article']}, Кол-во: {layer['count']}, Слой: {layer['layer']}"
                self.result_text.insert(tk.END, line + "\n")
                output.append(line)
            weight_line = f"  Общий вес: {pallet['weight']} кг\n"
            self.result_text.insert(tk.END, weight_line + "\n")
            output.append(weight_line)
    
        try:
            with open("результат_укладки.txt", "w", encoding="utf-8") as f:
                f.write("\n".join(output))
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{e}")
        else:
            messagebox.showinfo("Готово", "Результат сохранён в файл 'результат_укладки.txt'.")


if __name__ == "__main__":
    app = App()
    app.mainloop()

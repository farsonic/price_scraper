import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import csv
import time
import re
import os
import json
import webbrowser
import traceback
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError

# Import for Gemini AI
try:
    import google.generativeai as genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# Try to import optional libraries
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font
    HAS_EXCEL = True
except ImportError:
    HAS_EXCEL = False

class WoolworthsScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Woolworths Price Scraper")
        self.root.geometry("1200x700")

        # Data storage
        self.scraped_data = []
        self.is_scraping = False
        
        # Centralized settings variables
        self.api_key = tk.StringVar()
        self.model_var = tk.StringVar()
        
        # Updated model list based on current documentation
        self.available_models = [
            'gemini-2.5-flash', # Current model for speed and cost efficiency
            'gemini-2.5-pro'    # Current model for powerful, state-of-the-art performance
        ]

        # Default URLs
        self.default_urls = [
            'https://www.woolworths.com.au/shop/productdetails/160209/restor-concentrated-laundry-detergent-sheets-fresh-linen',
            'https://www.woolworths.com.au/shop/productdetails/164887/restor-concentrated-laundry-detergent-sheets-tropical',
            'https://www.woolworths.com.au/shop/productdetails/899864/undo-this-mess-laundry-detergent-sheets-spring-blossom',
            'https://www.woolworths.com.au/shop/productdetails/272748/earth-rescue-laundry-detergent-sheets-60-loads',
            'https://www.woolworths.com.au/shop/productdetails/909874/sheet-yeah-laundry-detergent-sheets-summer-daze',
            'https://www.woolworths.com.au/shop/productdetails/897214/dr-beckmann-laundry-detergent-sheets-universal',
            'https://www.woolworths.com.au/shop/productdetails/1122347429/cleaner-days-laundry-detergent-sheets-fragrance-free-60-pack',
            'https://www.woolworths.com.au/shop/productdetails/1122351339/cleaner-days-laundry-detergent-sheets-lemon-eucalyptus-60-pack',
            'https://www.woolworths.com.au/shop/productdetails/6021253/restor-dishwasher-detergent-sheets-lemon-burst',
            'https://www.woolworths.com.au/shop/productdetails/6020374/lucent-globe-dishwashing-detergent-sheets-fresh-lemon-scent',
            'https://www.woolworths.com.au/shop/productdetails/677834/earth-rescue-dishwasher-detergent-sheets-lemon-large',
            'https://www.woolworths.com.au/shop/productdetails/685913/earth-rescue-dishwasher-detergent-sheets-lemon-large',
            'https://www.woolworths.com.au/shop/productdetails/1122351535/cleaner-days-dishwasher-detergent-sheets-lemon-60-pack'
        ]

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=5, pady=5)
        self.urls_frame = ttk.Frame(notebook)
        notebook.add(self.urls_frame, text='URLs')
        self.setup_urls_tab()
        self.results_frame = ttk.Frame(notebook)
        notebook.add(self.results_frame, text='Results')
        self.setup_results_tab()
        self.log_frame = ttk.Frame(notebook)
        notebook.add(self.log_frame, text='Log')
        self.setup_log_tab()
        self.settings_frame = ttk.Frame(notebook)
        notebook.add(self.settings_frame, text='Settings')
        self.setup_settings_tab()
        self.setup_control_panel()

    # --- THIS IS THE CORRECTED FUNCTION ---
    def setup_urls_tab(self):
        # Frame for the listbox
        list_frame = ttk.Frame(self.urls_frame)
        list_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')
        self.url_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=15)
        self.url_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.url_listbox.yview)

        for url in self.default_urls:
            self.url_listbox.insert(tk.END, url)

        # A dedicated frame for all the control widgets below the list
        controls_frame = ttk.Frame(self.urls_frame)
        controls_frame.pack(fill='x', padx=10, pady=5)

        # Sub-frame for the URL entry field
        entry_frame = ttk.Frame(controls_frame)
        entry_frame.pack(fill='x', pady=(0, 5))
        
        self.url_entry = ttk.Entry(entry_frame)
        self.url_entry.pack(side='left', expand=True, fill='x', padx=(0, 5))
        
        ttk.Button(entry_frame, text="Add URL", command=self.add_url).pack(side='left')

        # Sub-frame for the action buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.pack(fill='x', expand=True)
        
        # Use grid inside this frame for stable, equal-width buttons
        button_frame.columnconfigure((0, 1, 2), weight=1)

        ttk.Button(button_frame, text="Remove Selected", command=self.remove_url).grid(row=0, column=0, sticky='ew', padx=2)
        ttk.Button(button_frame, text="Clear All", command=self.clear_urls).grid(row=0, column=1, sticky='ew', padx=2)
        ttk.Button(button_frame, text="Load Defaults", command=self.reset_urls).grid(row=0, column=2, sticky='ew', padx=2)


    def setup_results_tab(self):
        columns = ('Product', 'Price', 'Was', 'Unit Price', 'Promotion')
        self.tree = ttk.Treeview(self.results_frame, columns=columns, show='headings', height=20)
        self.tree.heading('Product', text='Product Name')
        self.tree.heading('Price', text='Current Price')
        self.tree.heading('Was', text='Was Price')
        self.tree.heading('Unit Price', text='Unit Price')
        self.tree.heading('Promotion', text='Promotion')
        self.tree.column('Product', width=400)
        self.tree.column('Price', width=100, anchor='center')
        self.tree.column('Was', width=100, anchor='center')
        self.tree.column('Unit Price', width=150, anchor='center')
        self.tree.column('Promotion', width=150, anchor='center')
        self.tree.bind("<Double-1>", self.on_item_double_click)
        vsb = ttk.Scrollbar(self.results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(column=0, row=0, sticky='nsew', padx=5, pady=5)
        vsb.grid(column=1, row=0, sticky='ns')
        hsb.grid(column=0, row=1, sticky='ew')
        self.results_frame.grid_columnconfigure(0, weight=1)
        self.results_frame.grid_rowconfigure(0, weight=1)
        self.summary_label = ttk.Label(self.results_frame, text="No data scraped yet")
        self.summary_label.grid(column=0, row=2, pady=5, sticky='w', padx=5)

    def setup_settings_tab(self):
        settings_pane = ttk.Frame(self.settings_frame, padding="10")
        settings_pane.pack(fill='x', expand=False, padx=10, pady=10)
        ttk.Label(settings_pane, text="Gemini API Key:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        api_key_instructions = "Get your API key from Google AI Studio. Settings are saved locally in a config.json file."
        ttk.Label(settings_pane, text=api_key_instructions, wraplength=500).grid(row=1, column=0, columnspan=2, sticky='w', pady=(0, 15))
        self.api_key_entry = ttk.Entry(settings_pane, textvariable=self.api_key, width=70, show='*')
        self.api_key_entry.grid(row=2, column=0, columnspan=2, sticky='ew', padx=(0, 5))
        ttk.Label(settings_pane, text="Gemini Model:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky='w', pady=(20, 5))
        model_instructions = "Select the AI model to use. 'Flash' is faster and cheaper, while 'Pro' is more powerful for complex analysis."
        ttk.Label(settings_pane, text=model_instructions, wraplength=500).grid(row=4, column=0, columnspan=2, sticky='w', pady=(0, 10))
        self.model_selector = ttk.Combobox(settings_pane, textvariable=self.model_var, values=self.available_models, state="readonly")
        self.model_selector.grid(row=5, column=0, columnspan=2, sticky='ew')
        self.model_selector.set(self.available_models[0])
        self.save_settings_button = ttk.Button(settings_pane, text="Save Settings", command=self.save_settings)
        self.save_settings_button.grid(row=6, column=0, columnspan=2, pady=(20, 0))
        settings_pane.grid_columnconfigure(0, weight=1)

    def setup_log_tab(self):
        self.log_text = scrolledtext.ScrolledText(self.log_frame, height=25, width=100, wrap=tk.WORD)
        self.log_text.pack(fill='both', expand=True, padx=10, pady=10)

    def setup_control_panel(self):
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill='x', side='bottom', padx=10, pady=10)
        self.progress = ttk.Progressbar(control_frame, mode='determinate')
        self.progress.pack(fill='x', padx=5, pady=(0, 10))
        options_frame = ttk.LabelFrame(control_frame, text="Options")
        options_frame.pack(side='left', padx=5)
        self.headless_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Headless Mode", variable=self.headless_var).pack(side='left', padx=5)
        self.debug_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Debug Mode", variable=self.debug_var).pack(side='left', padx=5)
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(side='right', padx=5)
        self.scrape_button = ttk.Button(buttons_frame, text="Start Scraping", command=self.start_scraping)
        self.scrape_button.pack(side='left', padx=5)
        self.csv_button = ttk.Button(buttons_frame, text="Export CSV", command=self.export_csv, state='disabled')
        self.csv_button.pack(side='left', padx=5)
        self.excel_button = ttk.Button(buttons_frame, text="Export Excel", command=self.export_excel, state='disabled')
        self.excel_button.pack(side='left', padx=5)
        self.ai_button = ttk.Button(buttons_frame, text="AI Analyse", command=self.start_ai_analysis, state='disabled')
        self.ai_button.pack(side='left', padx=5)

    def add_url(self):
        url = self.url_entry.get().strip()
        if url and url.startswith('https://www.woolworths.com.au/shop/productdetails/'):
            self.url_listbox.insert(tk.END, url)
            self.url_entry.delete(0, tk.END)
            self.log("Added URL: " + url)
        else:
            messagebox.showwarning("Invalid URL", "Please enter a valid Woolworths product URL")

    def remove_url(self):
        selection = self.url_listbox.curselection()
        if selection:
            self.url_listbox.delete(selection)
            self.log("Removed selected URL")

    def clear_urls(self):
        self.url_listbox.delete(0, tk.END)
        self.log("Cleared all URLs")

    def reset_urls(self):
        self.clear_urls()
        for url in self.default_urls:
            self.url_listbox.insert(tk.END, url)
        self.log("Reset to default URLs")

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def calculate_discount(self, current_price, was_price, promo_badge=""):
        if was_price and was_price not in ["Not applicable", "-", ""]:
            try:
                current_match = re.search(r'[\d.]+', str(current_price))
                was_match = re.search(r'[\d.]+', str(was_price))
                if current_match and was_match:
                    current, was = float(current_match.group()), float(was_match.group())
                    if was > current:
                        discount_amt, discount_pct = was - current, ((was - current) / was) * 100
                        if "1/2 Price" in promo_badge or 48 <= discount_pct <= 52: return discount_pct, "HALF PRICE!"
                        if "Special" in promo_badge: return discount_pct, f"{discount_pct:.0f}% OFF"
                        if discount_pct >= 30: return discount_pct, f"{discount_pct:.0f}% OFF!"
                        if discount_pct >= 20: return discount_pct, f"{discount_pct:.0f}% OFF"
                        if discount_pct > 0: return discount_pct, f"Save ${discount_amt:.2f}"
            except Exception: pass
        if promo_badge:
            if "1/2 Price" in promo_badge: return 50.0, "HALF PRICE!"
            if "Special" in promo_badge: return None, "SPECIAL"
        return None, ""

    def scrape_product_page(self, page, url):
        MAX_RETRIES = 3
        for attempt in range(MAX_RETRIES):
            try:
                page.goto(url, wait_until='domcontentloaded', timeout=30000)
                product_panel_selector = 'section[class*="product-details-panel_component_product-panel"]'
                page.wait_for_selector(product_panel_selector, timeout=20000)
                panel = page.locator(product_panel_selector)
                name = panel.locator('h1[class*="product-title_component_product-title"]').inner_text()
                price, was_price, cup_price, promo_badge = "Not found", "Not applicable", "Not found", ""
                try: price = panel.locator('div[class*="product-price_component_price-lead"]').inner_text(timeout=5000).replace('$', '').strip()
                except Exception: pass
                try: 
                    was_price_full = panel.locator('div[class*="product-unit-price_component_price-was"]').inner_text(timeout=2000)
                    was_price = re.search(r'[\d.]+', was_price_full).group()
                except Exception: pass
                try: cup_price = panel.locator('div[class*="product-unit-price_component_price-cup-string"]').inner_text(timeout=5000)
                except Exception: pass
                try: promo_badge = panel.locator('div[class*="product-stamp_message"]').inner_text(timeout=1000)
                except Exception: pass
                return {'name': name, 'price': price, 'was_price': was_price, 'cup_price': cup_price, 'url': url, 'promo_badge': promo_badge}
            except Exception as e:
                if attempt < MAX_RETRIES - 1:
                    time.sleep(2)
                    continue
                return {'error': str(e), 'url': url}

    def scraping_thread(self):
        urls = list(self.url_listbox.get(0, tk.END))
        if not urls:
            self.log("No URLs to scrape!")
            self.is_scraping = False
            self.scrape_button.config(text="Start Scraping", state='normal')
            return
        self.scraped_data, self.progress['value'] = [], 0
        self.tree.delete(*self.tree.get_children())
        self.progress['maximum'] = len(urls)
        self.log("Starting scraper...")
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=self.headless_var.get())
                context = browser.new_context(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36')
                page = context.new_page()
                try: page.goto("https://www.woolworths.com.au/", wait_until='domcontentloaded', timeout=30000)
                except Exception: self.log("Could not prime session, continuing anyway.")
                for i, url in enumerate(urls, 1):
                    self.log(f"Scraping {i}/{len(urls)}: {url.split('/')[-1]}")
                    data = self.scrape_product_page(page, url)
                    self.scraped_data.append(data)
                    if 'error' not in data:
                        _, promo_type = self.calculate_discount(data['price'], data['was_price'], data.get('promo_badge', ''))
                        price_display = f"${data['price']}" if data['price'] != "Not found" else "N/A"
                        was_display = f"${data['was_price']}" if data['was_price'] != "Not applicable" else "-"
                        self.tree.insert('', tk.END, values=(data['name'], price_display, was_display, data['cup_price'], promo_type or ""), tags=(data['url'],))
                        self.log(f"  ✓ {data['name']}")
                    else: self.log(f"  ✗ Error: {data['error']}")
                    self.progress['value'] = i
                    self.root.update_idletasks()
                browser.close()
        except Exception as e: self.log(f"An unexpected error occurred: {str(e)}")
        successful = sum(1 for d in self.scraped_data if 'error' not in d)
        promo_count = sum(1 for d in self.scraped_data if 'error' not in d and d.get('was_price') not in ["Not applicable", "-", ""])
        self.summary_label.config(text=f"Scraped: {successful}/{len(urls)} products | {promo_count} on special")
        if successful > 0:
            self.csv_button.config(state='normal')
            if HAS_EXCEL: self.excel_button.config(state='normal')
            if HAS_GEMINI: self.ai_button.config(state='normal')
        self.log("Scraping complete!")
        self.is_scraping = False
        self.scrape_button.config(text="Start Scraping", state='normal')

    def on_item_double_click(self, event):
        item_id = self.tree.identify_row(event.y)
        if item_id:
            tags = self.tree.item(item_id, "tags")
            if tags and tags[0]:
                url = tags[0]
                webbrowser.open(url)
                self.log(f"Opened URL in browser: {url}")

    def start_ai_analysis(self):
        if not HAS_GEMINI:
            messagebox.showerror("Gemini Not Available", "The 'google-generativeai' library is not installed.\nPlease run: pip install google-generativeai")
            return
        if not self.api_key.get():
            messagebox.showerror("API Key Missing", "Please go to the Settings tab and enter your Gemini API key.")
            return
        ai_window = tk.Toplevel(self.root)
        ai_window.title(f"Gemini AI Analysis ({self.model_var.get()})")
        ai_window.geometry("800x600")
        results_text = scrolledtext.ScrolledText(ai_window, height=20, width=80, wrap=tk.WORD, font=("Arial", 10))
        results_text.pack(padx=10, pady=10, fill='both', expand=True)
        results_text.insert('1.0', "Preparing data and contacting Gemini API... Please wait.")
        threading.Thread(target=self.run_gemini_analysis_thread, args=(results_text,), daemon=True).start()

    def run_gemini_analysis_thread(self, results_text_widget):
        model_name = ""
        try:
            genai.configure(api_key=self.api_key.get())
            model_name = self.model_var.get()
            model = genai.GenerativeModel(model_name)
            data_str = "Product Pricing Data:\n" + "-"*25 + "\n"
            for item in self.scraped_data:
                if 'error' not in item:
                    _, promo_type = self.calculate_discount(item['price'], item['was_price'], item.get('promo_badge', ''))
                    data_str += f"\nProduct: {item['name']}\n"
                    data_str += f"  Current Price: ${item['price']}\n"
                    if item['was_price'] != "Not applicable":
                        data_str += f"  Was Price: ${item['was_price']}\n"
                        data_str += f"  Promotion: {promo_type}\n"
                    data_str += f"  Unit Price: {item['cup_price']}\n"
            prompt = f"""You are a market analyst for a consumer goods company. Analyze the following competitive pricing data for laundry and dishwasher detergent sheets from Woolworths Australia.
**Data:**
{data_str}
**Your Task:**
Please provide a concise but insightful analysis covering these points:
1. **Overall Market Summary:** What is the general price range? Is the market competitive?
2. **Best Value:** Based on unit price (e.g., price per sheet or load), which products currently offer the best value for money?
3. **Promotion Analysis:** Which products are on special? What are the most common types of discounts (e.g., half price, percentage off)?
4. **Key Competitors:** Identify 2-3 key competitors based on their pricing and promotional activity.
5. **Strategic Recommendations:** Provide one key recommendation for a new product entering this market. Should it compete on price, quality, or another factor?
Structure your response with clear headings for each point. Be professional and data-driven."""
            if self.debug_var.get():
                self.log("--- DEBUG: AI PROMPT ---\n" + prompt + "\n--- END AI PROMPT ---")
            response = model.generate_content(prompt)
            results_text_widget.delete('1.0', tk.END)
            results_text_widget.insert('1.0', response.text)
        except Exception as e:
            error_message = f"An error occurred during AI analysis:\n\n{str(e)}"
            results_text_widget.delete('1.0', tk.END)
            results_text_widget.insert('1.0', error_message)
            if self.debug_var.get():
                self.log(f"--- DEBUG: AI ANALYSIS FAILED ---\nModel used: {model_name}\nError Type: {type(e).__name__}\nFull Traceback:\n{traceback.format_exc()}--- END DEBUG ---")

    def save_settings(self):
        key, model = self.api_key.get().strip(), self.model_var.get()
        if not key:
            messagebox.showwarning("Empty Key", "API key field is empty.")
            return
        try:
            with open("config.json", "w") as f:
                json.dump({"api_key": key, "model_name": model}, f, indent=4)
            self.log("Settings saved successfully to config.json.")
            messagebox.showinfo("Success", "Settings saved successfully.")
        except Exception as e:
            self.log(f"Error saving settings: {e}")
            messagebox.showerror("Error", f"Could not save settings to config.json:\n{e}")

    def load_settings(self):
        try:
            if os.path.exists("config.json"):
                with open("config.json", "r") as f:
                    settings = json.load(f)
                    self.api_key.set(settings.get("api_key", ""))
                    saved_model = settings.get("model_name", self.available_models[0])
                    self.model_var.set(saved_model if saved_model in self.available_models else self.available_models[0])
                    if self.api_key.get(): self.log("Loaded settings from config.json")
        except Exception as e:
            self.log(f"Could not load settings: {e}")
            self.api_key.set("")
            self.model_var.set(self.available_models[0])

    def start_scraping(self):
        if self.is_scraping: return
        self.is_scraping = True
        for button in [self.scrape_button, self.csv_button, self.excel_button, self.ai_button]:
            button.config(state='disabled')
        self.scrape_button.config(text="Scraping...")
        threading.Thread(target=self.scraping_thread, daemon=True).start()

    def export_csv(self):
        if not self.scraped_data: return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"woolworths_products_{timestamp}.csv"
        )
        if not filename: return
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['Product Name', 'Current Price', 'Was Price', 'Unit Price', 'Promotion', 'URL'])
                writer.writeheader()
                for item in (d for d in self.scraped_data if 'error' not in d):
                    _, promo_type = self.calculate_discount(item['price'], item['was_price'], item.get('promo_badge', ''))
                    writer.writerow({
                        'Product Name': item['name'], 
                        'Current Price': item['price'], 
                        'Was Price': item['was_price'], 
                        'Unit Price': item['cup_price'], 
                        'Promotion': promo_type, 
                        'URL': item['url']
                    })
            self.log(f"CSV exported to {filename}")
            messagebox.showinfo("Success", f"Data exported to {filename}")
        except Exception as e: 
            messagebox.showerror("Error", f"Failed to export CSV: {e}")

    def export_excel(self):
        if not HAS_EXCEL:
            messagebox.showwarning("Excel Not Available", "Please install openpyxl:\npip install openpyxl")
            return
        if not self.scraped_data: return
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile=f"woolworths_products_{timestamp}.xlsx"
        )
        if not filename: return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Woolworths Products"
            ws.append(['Product Name', 'Current Price', 'Was Price', 'Unit Price', 'Promotion', 'URL'])
            header_font, header_fill = Font(bold=True, color="FFFFFF"), PatternFill(start_color="366092", fill_type="solid")
            for cell in ws[1]: cell.font, cell.fill = header_font, header_fill
            half_price_fill, special_fill = PatternFill(start_color="FFC7CE", fill_type="solid"), PatternFill(start_color="FFEB9C", fill_type="solid")
            for item in (d for d in self.scraped_data if 'error' not in d):
                _, promo_type = self.calculate_discount(item['price'], item['was_price'], item.get('promo_badge', ''))
                try: price = float(item['price']) 
                except (ValueError, TypeError): price = ""
                try: was_price = float(item['was_price'])
                except (ValueError, TypeError): was_price = ""
                ws.append([item['name'], price, was_price, item['cup_price'], promo_type, item['url']])
                if "HALF PRICE" in promo_type.upper():
                    for cell in ws[ws.max_row]: cell.fill = half_price_fill
                elif promo_type:
                    for cell in ws[ws.max_row]: cell.fill = special_fill
            for column in ws.columns:
                max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=0)
                ws.column_dimensions[column[0].column_letter].width = min((max_length + 2), 60)
            wb.save(filename)
            self.log(f"Excel file exported to {filename}")
            messagebox.showinfo("Success", f"Data exported to {filename}")
        except Exception as e: messagebox.showerror("Error", f"Failed to export Excel: {e}")

def main():
    root = tk.Tk()
    app = WoolworthsScraperGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

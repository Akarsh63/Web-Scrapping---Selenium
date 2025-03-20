import time
import pandas as pd
import re
import os
import csv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from multiprocessing import Pool
import random
import openpyxl

# Define column order globally
COLUMN_ORDER = [
    "Product Url",
    "Product Name",
    "Description",
    "MRP",
    "Discounted Price",
    "Discount",
    "Highlights",
    "Images",
    "Variants",
    "Current Variant"
]

def create_driver(max_retries=3):
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--log-level=3")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")
    options.add_argument("--blink-settings=imagesEnabled=false")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    for attempt in range(max_retries):
        try:
            service = Service(ChromeDriverManager().install())
            service.log_path = "nul"
            time.sleep(1)
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except Exception as e:
            print(f"Failed to create driver (attempt {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                raise
            time.sleep(2)
    raise Exception("Failed to create Chrome driver after maximum retries")

def clean_image_url(image_url):
    match = re.match(r"(https://onemg\.gumlet\.io/)(?:[^/]+/)*([^/]+\.(?:jpg|png|jpeg|gif))", image_url)
    if match:
        base_url = match.group(1)
        file_name = match.group(2)
        return f"{base_url}{file_name}"
    return image_url

def append_to_temp_csv(data, temp_filename):
    """Append a single row of data to a temporary CSV file."""
    file_exists = os.path.isfile(temp_filename)
    with open(temp_filename, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=COLUMN_ORDER)
        if not file_exists:
            writer.writeheader()  # Write header if file doesn't exist
        # Convert lists and dictionaries to strings for CSV compatibility
        row_data = {}
        for col in COLUMN_ORDER:
            value = data.get(col, "")
            if col in ["Highlights", "Images", "Variants", "Current Variant"]:
                value = str(value)
            row_data[col] = value
        writer.writerow(row_data)
    print(f"Appended data for {data['Product Url']} to {temp_filename}")

def combine_temp_files(temp_files, final_filename, sheet_name):
    """Combine all temporary CSV files into a single Excel file."""
    all_data = []
    for temp_file in temp_files:
        if os.path.exists(temp_file):
            df = pd.read_csv(temp_file)
            all_data.append(df)
            # os.remove(temp_file)
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
        with pd.ExcelWriter(final_filename, engine='openpyxl') as writer:
            combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Combined all data into {final_filename}")
    else:
        print("No data to combine into Excel file.")

def extract_product_links(category_url):
    driver = create_driver()
    try:
        driver.get(category_url)
        try:
            cancel_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "UpdateCityModal__cancel-btn___2jWwS"))
            )
            cancel_button.click()
            print(f"Popup dismissed successfully for {category_url}")
        except Exception as e:
            print(f"No popup found or error dismissing it for {category_url}: {e}")

        product_data = []
        products = driver.find_elements(By.CLASS_NAME, "style__product-box___liepi")
        print(f"Found {len(products)} product boxes for {category_url}")

        for product in products:
            try:
                product_name = product.find_element(By.CLASS_NAME, "style__pro-title___2QwJy").text
                link_element = product.find_element(By.CLASS_NAME, "style__product-link___UB_67")
                product_link = link_element.get_attribute("href")
                description = product.find_element(By.CLASS_NAME, 'style__pack-size___2JQG7').text
                product_data.append({"Product Name": product_name, "Description": description, "Product Link": product_link})
            except Exception as e:
                print(f"Error extracting link: {e}")

        return product_data
    finally:
        driver.quit()

def scrape_product(args):
    product_url, description, temp_filename, sheet_name = args
    driver = create_driver()
    try:
        driver.get(product_url)
        try:
            cancel_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "UpdateCityModal__cancel-btn___2jWwS"))
            )
            cancel_button.click()
            print(f"Popup dismissed successfully for {product_url}")
        except Exception as e:
            print(f"No popup found or error dismissing it for {product_url}: {e}")

        product_data = driver.find_element(By.CLASS_NAME, 'OtcPage__top-container___2JKJ-')

        image_urls = []
        try:
            thumbnail_section = product_data.find_element(By.CLASS_NAME, "ProductImage__thumbnail-array___15c5x")
            images_section = thumbnail_section.find_elements(By.TAG_NAME, 'img')
            for section in images_section:
                raw_url = section.get_attribute("src")
                cleaned_url = clean_image_url(raw_url)
                image_urls.append(cleaned_url)
        except:
            image_urls = []

        try:
            product_name = product_data.find_element(By.CLASS_NAME, "ProductTitle__product-title___3QMYH").text
        except:
            product_name = "N/A"

        highlights = []
        try:
            highlight_section = product_data.find_element(By.CLASS_NAME, "ProductHighlights__highlights-text___dc-WQ")
            ul_elements = highlight_section.find_elements(By.TAG_NAME, "ul")
            for ul in ul_elements:
                for li in ul.find_elements(By.TAG_NAME, "li"):
                    highlights.append(li.text.strip())
        except:
            highlights = []

        price = "N/A"
        mrp = "N/A"
        discount = "N/A"
        try:
            price_box = product_data.find_element(By.CLASS_NAME, "OtcPriceBox__price-box___p13HY")
            price_elements = price_box.find_elements(By.XPATH, ".//*[contains(text(), '₹')]")
            prices = []
            for elem in price_elements:
                text = elem.text.strip()
                if text and "₹" in text and "% off" not in text:
                    prices.append((elem, text))

            if len(prices) >= 2:
                for i, (elem, text) in enumerate(prices):
                    is_strikethrough = (
                        elem.tag_name in ["strike", "s"] or
                        bool(elem.find_elements(By.TAG_NAME, "strike")) or
                        bool(elem.find_elements(By.TAG_NAME, "s")) or
                        "line-through" in elem.value_of_css_property("text-decoration")
                    )
                    if not is_strikethrough:
                        child_elements = elem.find_elements(By.XPATH, ".//*")
                        for child in child_elements:
                            if "line-through" in child.value_of_css_property("text-decoration"):
                                is_strikethrough = True
                                break

                    is_mrp_label = "MRP" in text.upper()
                    
                    if is_strikethrough or is_mrp_label:
                        mrp = text
                        price = prices[1 - i][1]
                        break
                else:
                    mrp = prices[0][1]
                    price = prices[1][1]
            elif len(prices) == 1:
                price = prices[0][1]
                mrp = price

            if price != "N/A":
                price_match = re.search(r'\d+\.?\d*', price)
                price = price_match.group() if price_match else "N/A"
            if mrp != "N/A":
                mrp_match = re.search(r'\d+\.?\d*', mrp)
                mrp = mrp_match.group() if mrp_match else "N/A"

            discount_elements = price_box.find_elements(By.XPATH, ".//*[contains(text(), '% off')]")
            for elem in discount_elements:
                text = elem.text.strip()
                if text and "% off" in text:
                    discount = text
                    break

        except Exception as e:
            print(f"Error extracting price details for {product_url}: {e}")

        variants = []
        current_variant = []
        try:
            variants_containers = product_data.find_elements(By.CLASS_NAME, "OtcVariants__container___2Y3D2")
            for variant_container in variants_containers:
                header = variant_container.find_element(By.CLASS_NAME, "OtcVariants__header___2q6Sa")
                variant_name = header.find_element(By.TAG_NAME, "h3").text
                variant_name = re.sub(r'\s*\(\d+\)', '', variant_name).strip()
                curr_container = {variant_name: []}

                variant_divs = variant_container.find_elements(By.CLASS_NAME, "OtcVariants__variant-div___2l321")
                variant_values = []
                for variant_div in variant_divs:
                    variant_item = variant_div.find_element(By.CLASS_NAME, "OtcVariantsItem__container___2ldJL")
                    variant_text = variant_item.find_element(By.CLASS_NAME, "OtcVariantsItem__variant-text___1Grsz").text.strip()
                    variant_values.append(variant_text)

                    if "OtcVariantsItem__selected___1wDpJ" in variant_item.get_attribute("class"):
                        selected_value = variant_text
                        current_variant.append({variant_name: selected_value})

                curr_container[variant_name] = variant_values
                variants.append(curr_container)

        except Exception as e:
            print(f"No variants found for {product_url}: {e}")
            variants = []
            current_variant = []

        time.sleep(random.uniform(0.5, 1.5))

        product_details = {
            "Product Url": product_url,
            "Product Name": product_name,
            "Description": description if description else "N/A",
            "MRP": mrp,
            "Discounted Price": price,
            "Discount": discount,
            "Highlights": highlights,
            "Images": image_urls,
            "Variants": variants,
            "Current Variant": current_variant
        }
        append_to_temp_csv(product_details, temp_filename)

        return product_details

    except Exception as e:
        print(f"Error scraping product {product_url}: {e}")
        product_details = {
            "Product Url": product_url,
            "Product Name": "Error",
            "Description": description if description else "N/A",
            "MRP": "N/A",
            "Discounted Price": "N/A",
            "Discount": "N/A",
            "Highlights": [],
            "Images": [],
            "Variants": [],
            "Current Variant": []
        }
        append_to_temp_csv(product_details, temp_filename)
        return product_details

    finally:
        driver.quit()

def main(categories):
    category_data = {}

    for category in categories:
        category_name, base_url = category[0], category[1]
        print(f"Starting extraction for category: {category_name}")

        all_product_urls = []
        current_url = base_url
        page_number = 1

        while True:
            print(f"Extracting links for category: {category_name}, page {page_number} ({current_url})")
            product_urls = extract_product_links(current_url)
            all_product_urls.extend(product_urls)

            driver = create_driver()
            try:
                driver.get(current_url)
                next_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "next"))
                )
                if "link-disabled" in next_button.get_attribute("class"):
                    print("No more pages to scrape.")
                    break

                next_button.click()
                time.sleep(2)
                current_url = driver.current_url
                page_number += 1
            except Exception as e:
                print(f"No 'Next' button found or error navigating to next page: {e}")
                break
            finally:
                driver.quit()

        category_data[category_name] = pd.DataFrame(all_product_urls)
        print(f"Total products extracted for {category_name}: {len(all_product_urls)}")

        with pd.ExcelWriter("product_links.xlsx") as writer:
            for category, df in category_data.items():
                df.to_excel(writer, sheet_name=category, index=False)

def main_2(categories, num_processes=4):
    for category in categories:
        category_name = category[0]
        print(f"Scraping product details for category: {category_name}")

        try:
            df = pd.read_excel("product_links.xlsx", sheet_name=category_name)
            product_urls = df["Product Link"].tolist()
            descriptions = df["Description"].tolist()
            print(f"Found {len(product_urls)} product URLs for {category_name}")
        except Exception as e:
            print(f"Error reading product links for {category_name}: {e}")
            continue

        # Create a list to store temporary CSV filenames
        temp_files = []
        filename = f"{category_name}.xlsx"
        # Distribute URLs across processes
        chunk_size = len(product_urls) // num_processes if num_processes > 0 else len(product_urls)
        chunks = [product_urls[i:i + chunk_size] for i in range(0, len(product_urls), chunk_size)]
        desc_chunks = [descriptions[i:i + chunk_size] for i in range(0, len(descriptions), chunk_size)]

        # Create arguments for each process
        args = []
        for i, (url_chunk, desc_chunk) in enumerate(zip(chunks, desc_chunks)):
            temp_filename = f"temp_{category_name}_{i}.csv"
            temp_files.append(temp_filename)
            for url, desc in zip(url_chunk, desc_chunk):
                args.append((url, desc, temp_filename, category_name))

        # Use multiprocessing to scrape products in parallel
        with Pool(processes=num_processes) as pool:
            pool.map(scrape_product, args)

        # Combine temporary CSV files into a single Excel file
        combine_temp_files(temp_files, filename, category_name)

    print(f"Finished scraping product details. Saved to {category_name}.xlsx")

if __name__ == "__main__":
    categories = [
        ['Pet Care', 'https://www.1mg.com/categories/pet-care-612'],
        # Add other categories as needed
    ]

    # main(categories)  # Uncomment to extract product links
    main_2(categories, num_processes=4)

import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from multiprocessing import Pool
import random

def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # Run in headless mode
    options.add_argument("--log-level=3")  # Suppress logs
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/134.0.0.0 Safari/537.36")
    # Disable images to speed up page loads
    options.add_argument("--blink-settings=imagesEnabled=false")
    service = Service(ChromeDriverManager().install())
    service.log_path = "nul"  # On Windows, "nul" discards logs
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def clean_image_url(image_url):
    match = re.match(r"(https://onemg\.gumlet\.io/)(?:[^/]+/)*([^/]+\.(?:jpg|png|jpeg|gif))", image_url)
    if match:
        base_url = match.group(1)  
        file_name = match.group(2) 
        return f"{base_url}{file_name}"
    return image_url

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

def scrape_product(product_url):
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
        
        # Extract variants
        variants = []
        current_variant = []
        try:
            # Find all variant containers (one for each variant type, e.g., "Potency", "Pack Size")
            variants_containers = product_data.find_elements(By.CLASS_NAME, "OtcVariants__container___2Y3D2")

            for variant_container in variants_containers:
                # Extract variant name (e.g., "Potency")
                header = variant_container.find_element(By.CLASS_NAME, "OtcVariants__header___2q6Sa")
                variant_name = header.find_element(By.TAG_NAME, "h3").text
                variant_name = re.sub(r'\s*\(\d+\)', '', variant_name).strip()
                curr_container = {variant_name: []}

                # Find all variant value containers (nested inside an extra <div class="">)
                variant_divs = variant_container.find_elements(By.CLASS_NAME, "OtcVariants__variant-div___2l321")

                # Extract all variant values
                variant_values = []
                for variant_div in variant_divs:
                    variant_item = variant_div.find_element(By.CLASS_NAME, "OtcVariantsItem__container___2ldJL")
                    # Extract the text from the inner variant-text div to avoid extra whitespace/comments
                    variant_text = variant_item.find_element(By.CLASS_NAME, "OtcVariantsItem__variant-text___1Grsz").text.strip()
                    variant_values.append(variant_text)

                    # Check if this is the selected variant
                    if "OtcVariantsItem__selected___1wDpJ" in variant_item.get_attribute("class"):
                        selected_value = variant_text
                        current_variant.append({variant_name: selected_value})

                # Store the variant values in the current container
                curr_container[variant_name] = variant_values

                # Append the current container to the variants list
                variants.append(curr_container)

        except Exception as e:
            print(f"No variants found for {product_url}: {e}")
            variants = []
            current_variant = []

        time.sleep(random.uniform(0.5, 1.5))

        return {
            "Product Url": product_url,
            "Product Name": product_name,
            "MRP": mrp,
            "Discounted Price": price,
            "Discount": discount,
            "Highlights": highlights,
            "Images": image_urls,
            "Variants": variants,
            "Current Variant": current_variant
        }
    except Exception as e:
        print(f"Error scraping product {product_url}: {e}")
        return {
            "Product Url": product_url,
            "Product Name": "Error",
            "MRP": "N/A",
            "Discounted Price": "N/A",
            "Discount": "N/A",
            "Highlights": [],
            "Images": [],
            "Variants": [],
            "Current Variant": []
        }
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
                time.sleep(2)  # Reduced delay
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
    detailed_data = {}

    for category in categories:
        category_name = category[0]
        print(f"Scraping product details for category: {category_name}")

        try:
            df = pd.read_excel(f"product_links.xlsx", sheet_name=category_name)
            product_urls = df["Product Link"].tolist()  # Adjust as needed
            descriptions = df["Description"].tolist()  # Get descriptions
            print(f"Found {len(product_urls)} product URLs for {category_name}")
        except Exception as e:
            print(f"Error reading product links for {category_name}: {e}")
            continue

        # Use multiprocessing to scrape products in parallel
        with Pool(processes=num_processes) as pool:
            product_details_list = pool.map(scrape_product, product_urls)
        
        for i, product_details in enumerate(product_details_list):
            product_details["Description"] = descriptions[i] if i < len(descriptions) else "N/A"

        column_order = [
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

        # Create DataFrame with specified column order
        detailed_data[category_name] = pd.DataFrame(product_details_list)[column_order]

    with pd.ExcelWriter(f"{category_name}.xlsx") as writer:
        for category, df in detailed_data.items():
            df.to_excel(writer, sheet_name=category, index=False)

    print(f"Finished scraping product details. Saved to {category_name}.xlsx")

if __name__ == "__main__":
    categories = [
        ['Diabetes', 'https://www.1mg.com/categories/diabetes-1'],
        # ['Vitamins & Nutrition', 'https://www.1mg.com/categories/vitamins-nutrition-5'],
        # ['Nutritional Drinks', 'https://www.1mg.com/categories/nutritional-drinks-196'],
        # ['Pet Care', 'https://www.1mg.com/categories/pet-care-612'],
        # ['Sexual Wellness', 'https://www.1mg.com/categories/sexual-wellness-22'],
        # ['Stomach Care', 'https://www.1mg.com/categories/stomach-care-30'],
        # ['Fitness Supplements', 'https://www.1mg.com/categories/fitness-supplements-7'],
        # ['Pain Relief', 'https://www.1mg.com/categories/pain-relief-32'],
        # ['Herbal Juices', 'https://www.1mg.com/categories/herbal-juice-174'],
        # ['Healthy Snacks', 'https://www.1mg.com/categories/healthy-snacks-623'],
        # ['Immunity Boosters', 'https://www.1mg.com/categories/immunity-boosters-142'],
        # ['Covid Essentials', 'https://www.1mg.com/categories/coronavirus-prevention-925'],
        # ['Personal Care', 'https://www.1mg.com/categories/personal-care-products-18'],
        # ['Supports & Braces', 'https://www.1mg.com/categories/supports-braces-15'],
        # ['Hair Care', 'https://www.1mg.com/categories/hair-skin-care-event-hair-care-3305'],
        # ['The Ayurveda Store', 'https://www.1mg.com/categories/ayurveda-products-45'],
        # ['Winter Care', 'https://www.1mg.com/categories/winter-care-essentials-3362'],
        # ['Elderly Care', 'https://www.1mg.com/categories/elderly-care-21'],
        # ['Rehydration Beverages', 'https://www.1mg.com/categories/rehydration-beverages-382'],
        # ['Health Conditions', 'https://www.1mg.com/categories/health-conditions-28'],
        # ['First Aid', 'https://www.1mg.com/categories/first-aid-36'],
        # ['Devices in Healthcare', 'https://www.1mg.com/categories/devices-healthcare-1288'],
        # ['Cold, Cough & Fever', 'https://www.1mg.com/categories/cold-cough-fever-37'],
        # ['Baby Care', 'https://www.1mg.com/categories/baby-care-24'],
        # ['Anti-Pollution', 'https://www.1mg.com/categories/anti-pollution-237'],
        # ['Fever & Headache', 'http://1mg.com/categories/health-conditions/fever-38'],
        # ['Skin Infection', 'https://www.1mg.com/categories/skin-infection-1184'],
        # ['Featured', 'https://www.1mg.com/categories/featured-128'],
        # ['Eye & Ear Care', 'https://www.1mg.com/categories/health-conditions/eye-ear-care-3063'],
        # ['Monitoring Devices', 'https://www.1mg.com/categories/monitoring-devices-1869'],
        # ['Healthcare Devices', 'https://www.1mg.com/categories/medical-devices-12'],
        # ['Tata Neu', 'https://www.1mg.com/categories/tata-neu-1305'],
        # ['Homeopathy','https://www.1mg.com/categories/homeopathy-57'],
        # ['Face Care', 'https://www.1mg.com/categories/hair-skin-care-event-face-care-3306'],
        # ['Body Care', 'https://www.1mg.com/categories/hair-skin-care-event-body-care-3307']
    ]

    
    # main(categories)  # Uncomment to extract product links
    main_2(categories, num_processes=4)
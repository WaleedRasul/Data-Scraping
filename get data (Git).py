import requests
from scrapy import Selector
import pandas as pd
from multiprocessing import Pool
import re

# Add comments to explain the purpose of cookies and headers
cookies = {
    # Insert cookies here
}

headers = {
    # Insert headers here
}


def get_data(url):
    print(url)

    asin = ''
    # Use regex to extract ASIN from the URL
    pattern = r'/(\d+)/product\.html'
    match = re.search(pattern, url)
    if match:
        asin = match.group(1)
        print(f'ASIN: {asin}')

    response = requests.get(url, cookies=cookies, headers=headers)

    if response.status_code == 200:
        res = Selector(text=response.text)   # Get the text of response by parsing it using scrapy 'selector' framework


        # Extract data using XPath expressions
        titles = res.xpath("//h1[@data-test='product-title']/text()").getall()
        full_title = ''.join(titles).strip()
        print(f'Title: {full_title}')

        brand = res.xpath("//span[@data-testid='brand-info']/a/text()").get()
        print(f'Brand: {brand}')

        price = res.xpath("//div[@data-testid='current-price']/div[@class='css-1olsk4d e1eyx97t2']/text()").get()
        if price:
            price_part = price.split('$')[1]
            if ',' in price_part:
                # Remove commas from the price part
                price_part = price_part.replace(',', '')
            try:
                price = float(price_part)
                print(f'Price: {price}')
            except IndexError:
                print("Error: Price format is not as expected.")
        else:
            print("Error: Price not found on the page.")

        reviews_count = res.xpath(
            "//button[@aria-label='Go to reviews section']//span[@data-testid='reviews-count']/text()").get()
        print(f'Reviews count: {reviews_count}')

        rating = res.xpath("//div[@class='css-1uxmhx9 e7xp0l512']//span[@class='css-xsq97o e7xp0l59']/text()").get()
        if rating and 'NaN' in rating:
            rating = rating.replace('NaN', '')
        print(f'Rating: {rating}')

        color = res.xpath(
            "//div[@id='spec-height']//div[@class='css-1qfex5m e20w0xa10'][preceding-sibling::div[contains(text(),'Color')]]/text()").get()
        print(f'Color: {color}')

        material = res.xpath(
            "//div[@id='spec-height']//div[@class='css-1qfex5m e20w0xa10'][preceding-sibling::div[contains(text(),'Material')]]/text()").get()
        print(f'Material: {material}')

        dimensions = res.xpath(
            "//div[@id='spec-height']//div[@class='css-1qfex5m e20w0xa10'][preceding-sibling::div[contains(text(),'Assembled Dimensions')]]/text()").get()
        if dimensions and ('See Description' in dimensions or 'See Details' in dimensions):
            dimensions = dimensions.replace('See Description', '')
            dimensions = dimensions.replace('See Details', '')
            print(f'Dimensions: {dimensions}')

        cushion_style = res.xpath(
            "//div[@id='spec-height']//div[@class='css-1qfex5m e20w0xa10'][preceding-sibling::div[contains(text(),'Cushion Type')]]/text()").get()
        print(f'Cushion Type: {cushion_style}')

        image_url = res.xpath("//div[@data-testid='magnifier wrapper']/div/img/@src").get()
        print(f'Image URL: {image_url}')

        xpaths = [
            "//div[@id='description-height']//div[preceding-sibling::div[contains(text(),'Details:')]]/p//text()",
            "//div[@id='description-height']//div[preceding-sibling::div[contains(text(),'Details:')]]/text()",
            "//div[@id='description-height']//div[preceding-sibling::div[contains(text(),'Details:')]]/ul/li/text()",
            "//div[@id='description-height']//div[preceding-sibling::div[contains(text(),'Details:')]]/p/strong/text(),"
            "//div[@id='description-height']//div[preceding-sibling::div[contains(text(),'Details:')]]/p[1]/text()",
        ]

        description = ''

        for xpath in xpaths:
            text_list = res.xpath(xpath).extract()
            # print(f"Result of XPath '{xpath}': {text}")
            if text_list:
                text = ' '.join(text_list)
                if text.strip():
                    description = text.strip()
                    break
        print(f'Description: {description}')

        # Make a Data Dictionary to assign the data their names
        data_dict = {
            'ASIN': asin,
            'Title': full_title,
            'Brand': brand if brand else '-',
            'Cushion Type': cushion_style if cushion_style else '-',
            'Manufacturer': '-',
            'Price': price if price else '-',
            'Rating': rating if rating else '0',
            'Reviews': reviews_count if reviews_count else '0',
            'Category Rank': '-',
            'Dimensions': dimensions if dimensions else '-',
            'Color': color if color else '-',
            'Material': material if material else '-',
            'Description': description,
            'URL': url,
            'Image URL': image_url if image_url else '-',
        }

    else:
        print("Error:", response.status_code)

    return data_dict


if __name__ == '__main__':
    # Read URLs from Excel file
    df = pd.read_excel("final testing variant urls.xlsx")
    urls = df["Completed URL"].tolist()
    print(len(urls))

    # Use multiprocessing to speed up data extraction
    p = Pool(25)  # 25 threads means 25 requests at a time
    results = p.map(get_data, urls)
    p.terminate()
    p.join()

    # Filter out the data from lists
    result_data = [result for result in results if result is not None]

    # Convert extracted data to DataFrame and save to Excel
    df = pd.DataFrame(result_data)
    df.to_excel('output.xlsx', index=False)

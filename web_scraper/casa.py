import requests, os
from bs4 import BeautifulSoup
import pandas as pd

def casa_webscrap() -> None:
    """
    Web scraper for website Civil Aviation Safety Authority.
    """
    path, url, df = str(os.path.abspath(os.getcwd())).replace('\\', '/') + '/', 'https://shop.casa.gov.au', list()

    session = requests.Session()
    html = session.get(f'{url}/collections').text
    c = 'product-grid product-grid--per-row-4 product-grid--per-row-mob-2 product-grid--image-shape-square'
    blocks = BeautifulSoup(html).find_all('div', attrs={'class': c})[0]
    for bs in blocks:
        bs = str(bs)
        if '/collections/' in bs:
            block, k = bs[bs.index('/collections/'):bs.index('" title="')].split('/')[-1].strip(), 1
            while True:
                field, pr_id, pr_title, pr_price, pr_link = session.get(f'{url}/collections/{block}?page={k}').text, '', '', '', ''
                product_block = BeautifulSoup(field).find_all('product-block', class_='product-block')
                for pr in product_block:
                    pr = str(pr).split('\n')
                    for s in pr:
                        if 'data-product-id' in s:
                            pr_id = s[s.index('-id="')+5:-2]
                        if 'product-block__title' in s:
                            pr_title = s[s.index('left">')+6:s.index('</div>')].replace('&amp;', '&')
                        if 'price__current' in s:
                            pr_price = s.split('$')[-1]
                        if '<a aria-hidden="true" class="product-link"' in s:
                            pr_link = s[s.index('href="')+6:s.index('" tabindex="')].strip()
                    df.append([pr_id, block, pr_title, pr_price, f'{url}{pr_link}'])
                k += 1
                if not pr_id:
                    break
    pd.DataFrame(df, columns=['Code', 'Collection', 'Name', 'Price', 'URL']).to_excel(f'{path}ShopCasa.xlsx', index=False)

if __name__ == '__main__':
    casa_webscrap()

import requests
from newspaper import Article
import httpx
import asyncio

async def scrape_site(url: str) -> str:
    async with httpx.AsyncClient() as client:
        response = await client.get(url)
        if response.status_code != 200:
            raise Exception(f"Failed to fetch {url}: {response.reason_phrase}")
        response.encoding = "utf-8"
        html = response.text
        return html

async def main():
    html = await scrape_site("https://www.prodyna.com/de/jobs/zurich-full-stack")

    print(html)
    # Get HTML from the real job page
    article = Article("mock")
    article.set_html(html)
    article.parse()

    print(article.title)
    print(article.text)


asyncio.run(main())
def extract_job_description_html(url):
    resp = requests.get(url)
    resp.raise_for_status()  # raise error if status != 200
    return resp.text

# Fuck this, I want to use TypeScript


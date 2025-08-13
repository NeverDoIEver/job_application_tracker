import requests
from newspaper import Article

def extract_job_description_json(url):
    # Convert the public job URL to the Workday API endpoint
    job_key = url.split("/")[-1]
    api_url = f"https://zuehlke.wd3.myworkdayjobs.com/wday/cxs/zuehlke/Zuhlke-Careers/job/Schlieren/{job_key}"

    resp = requests.get(api_url)
    data = resp.json()

    description_html = data.get("jobPostingInfo", {}).get("jobDescription", "")
    return description_html

html_name = "https://zuehlke.wd3.myworkdayjobs.com/en-US/Zuhlke-Careers/job/Schlieren/Data---AI-Consultant-for-Banking_JR100231?locations=02c489c2044f1003c6232f78e4b70000&locations=02c489c2044f1003c62342be5b3f0000"


article = Article('mock.ch')
article.set_html(extract_job_description_json(html_name))
article.parse()

print(article.title)
print(article.text)

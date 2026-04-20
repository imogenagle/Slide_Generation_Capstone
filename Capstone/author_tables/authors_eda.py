import pandas as pd
from pathlib import Path

BASE = Path(__file__).resolve().parent

authors = pd.read_csv(BASE / "authors.csv")
print(authors.head())
paper_authors = pd.read_csv(BASE / "paper_authors.csv")
print(paper_authors.head())
papers = pd.read_csv(BASE / "papers.csv")
print(papers.head())

print(authors["paper_count"].value_counts().reset_index(name="count"))

print(papers["paper_pdf_exists"].value_counts())

print(authors[authors["paper_count"] == 10])

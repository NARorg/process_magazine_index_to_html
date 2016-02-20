#!/usr/bin/env python

import sys
import openpyxl
import process_functions
import path

#want at least 4 html outputs
#chronological listing of every issue
#chrono listing of every article
#listing of every article in every category
#listing of every article by every author

#also want to create a new folder of PDFs with a tiny amount of 
#obfuscation in place to prevent someone from guessing every PDF name 
#and downloading them all (want to at least give people a challenge).

wb = openpyxl.load_workbook(sys.argv[1], guess_types=True)
ws = wb.active


rowCount = process_functions.count_rows(ws)
colCount = process_functions.count_cols(ws)
headers, dict = process_functions.make_dict(ws)
issuesToIndices = process_functions.gather_issues(headers, dict)
issueHtml = process_functions.output_issue_html(
    issuesToIndices.keys(), path.globalPath)
outIssueHtml = open('chronological_by_issue.html', 'w')
outIssueHtml.write(issueHtml)
outIssueHtml.close()
articleHtml = process_functions.output_article_html_unique(
    issuesToIndices.keys(),
    issuesToIndices, dict, path.globalPath)
outArticleHtml = open('chronological_by_article.html', 'w')
outArticleHtml.write(articleHtml)
outArticleHtml.close()
categoryHtml = process_functions.output_category_html(
    issuesToIndices.keys(), issuesToIndices, dict, path.globalPath)
outCategoryHtml = open('category.html', 'w')
outCategoryHtml.write(categoryHtml)
outCategoryHtml.close()
authorHtml = process_functions.output_author_html(
   issuesToIndices.keys(), issuesToIndices, dict, path.globalPath)
outAuthorHtml = open('author.html', 'w')
outAuthorHtml.write(authorHtml)
outAuthorHtml.close()


rows = process_functions.count_rows(ws)
print rows

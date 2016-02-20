#!/usr/bin/env python

import openpyxl
import unittest
import string
import collections
import sys
import operator

indexString = u'IDX'
authorString = u'AUTHOR'
sourceString = u'SOURCE'
volumeString = u'Vol.'
numberString = u'No.'
pageString = u'Page'
monthString = u'Month'
monthNumberString = u'# Month'
yearString = u'Year'
topicString = u'Topic'
titleString = u'Title'

testPath = 'http://fake.path/'

nameDict = {'MR': 'Model Rocketeer'}

goldIssueString = """<p><h1>1973</h1></p>\n<p><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocketeer Jan 1973 Volume 14 Number 1</a></p>\n<p><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Model Rocketeer Dec 1973 Volume 15 Number 11</a></p>\n"""

goldArticleString = """<p><h1>1973</h1></p>
<p><h2><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocketeer Jan 1973 Volume 14 Number 1</a></h2></p>
<ul><li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Editor's Nook, by Sadowski, Elaine, Page 4</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR In Action, by Unknown, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">First World Model Rocket Championships-VRSAC '72, by Pearson, Ed, Pages 6-7, 14-15</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Technical Feature The Effect of Delayed Staging on a Multi-staged Model Rocket's Performance, by Kuechler, Thomas, Pages 8-9</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR News, by Unknown, Pages 11, 14</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Unearthly-MARS VII Aberdeen Proving Ground, Maryland October 14-15, 1972, by Diller, Elisa, Page 12</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), by Newill, David, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">New From the Manufacturers, by Lieber, Robert, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), by Newill, David, Page 13</a></li>
</ul>
<p><h2><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Model Rocketeer Dec 1973 Volume 15 Number 11</a></h2></p>
<ul><li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Editor's Nook, by Sadowski, Elaine, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Roc Egglofter, by Cole, Gary, Page 9</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Section Highlights, by Blickenstaff, Jan, Page 10</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">R & D Summary A Polaroid Camroc System, by Griffith, Patrick M., Page 11</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, by Medina, Tony, Pages 12-13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, by Medina, Tony, Pages 12-13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Saturn V Launcher, The, by Gross, Paul, Pages 14-19</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">NAR In Action, by Wright, Ron, Page 20</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Oscar Two Stage Sport Rocket, by Conner II, Paul, Page 21</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Television and Model Rocketry, by Shenosky, Larry, Page 22</a></li>
</ul>
"""

goldArticleUniqueString = """<p><h1>1973</h1></p>
<p><h2><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocketeer Jan 1973 Volume 14 Number 1</a></h2></p>
<ul><li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Editor's Nook, by Sadowski, Elaine, Page 4</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR In Action, by Unknown, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">First World Model Rocket Championships-VRSAC '72, by Pearson, Ed, Pages 6-7, 14-15</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Technical Feature The Effect of Delayed Staging on a Multi-staged Model Rocket's Performance, by Kuechler, Thomas, Pages 8-9</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR News, by Unknown, Pages 11, 14</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Unearthly-MARS VII Aberdeen Proving Ground, Maryland October 14-15, 1972, by Diller, Elisa, Page 12</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), by Newill, David, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">New From the Manufacturers, by Lieber, Robert, Page 13</a></li>
</ul>
<p><h2><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Model Rocketeer Dec 1973 Volume 15 Number 11</a></h2></p>
<ul><li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Editor's Nook, by Sadowski, Elaine, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Roc Egglofter, by Cole, Gary, Page 9</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Section Highlights, by Blickenstaff, Jan, Page 10</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">R & D Summary A Polaroid Camroc System, by Griffith, Patrick M., Page 11</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, by Medina, Tony, Pages 12-13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Saturn V Launcher, The, by Gross, Paul, Pages 14-19</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">NAR In Action, by Wright, Ron, Page 20</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Oscar Two Stage Sport Rocket, by Conner II, Paul, Page 21</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Television and Model Rocketry, by Shenosky, Larry, Page 22</a></li>
</ul>
"""


goldCategoryString = """<p><h1>Columns</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Editor's Nook, by Sadowski, Elaine, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 4</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR In Action, by Unknown, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR News, by Unknown, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 11, 14</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Editor's Nook, by Sadowski, Elaine, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Section Highlights, by Blickenstaff, Jan, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 10</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">NAR In Action, by Wright, Ron, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 20</a></li>
</ul>
<p><h1>Competition - Boost Glider</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), by Newill, David, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, by Medina, Tony, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 12-13</a></li>
</ul>
<p><h1>Competition - Duration</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Roc Egglofter, by Cole, Gary, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 9</a></li>
</ul>
<p><h1>Competition - Rocket Glider</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), by Newill, David, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, by Medina, Tony, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 12-13</a></li>
</ul>
<p><h1>Construction</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Oscar Two Stage Sport Rocket, by Conner II, Paul, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 21</a></li>
</ul>
<p><h1>Contest Reports</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Unearthly-MARS VII Aberdeen Proving Ground, Maryland October 14-15, 1972, by Diller, Elisa, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 12</a></li>
</ul>
<p><h1>International Spacemodeling</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">First World Model Rocket Championships-VRSAC '72, by Pearson, Ed, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 6-7, 14-15</a></li>
</ul>
<p><h1>Manufacturers</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">New From the Manufacturers, by Lieber, Robert, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
</ul>
<p><h1>Other</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Television and Model Rocketry, by Shenosky, Larry, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 22</a></li>
</ul>
<p><h1>Photography and Video</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">R & D Summary A Polaroid Camroc System, by Griffith, Patrick M., in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 11</a></li>
</ul>
<p><h1>Space History</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Saturn V Launcher, The, by Gross, Paul, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 14-19</a></li>
</ul>
<p><h1>Technical Articles</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Technical Feature The Effect of Delayed Staging on a Multi-staged Model Rocket's Performance, by Kuechler, Thomas, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 8-9</a></li>
</ul>
"""

goldAuthorString = """<p><h1>Blickenstaff, Jan</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Section Highlights, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 10</a></li>
</ul>
<p><h1>Cole, Gary</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Roc Egglofter, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 9</a></li>
</ul>
<p><h1>Conner II, Paul</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Oscar Two Stage Sport Rocket, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 21</a></li>
</ul>
<p><h1>Diller, Elisa</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Unearthly-MARS VII Aberdeen Proving Ground, Maryland October 14-15, 1972, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 12</a></li>
</ul>
<p><h1>Griffith, Patrick M.</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">R & D Summary A Polaroid Camroc System, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 11</a></li>
</ul>
<p><h1>Gross, Paul</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Saturn V Launcher, The, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 14-19</a></li>
</ul>
<p><h1>Kuechler, Thomas</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Technical Feature The Effect of Delayed Staging on a Multi-staged Model Rocket's Performance, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 8-9</a></li>
</ul>
<p><h1>Lieber, Robert</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">New From the Manufacturers, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
</ul>
<p><h1>Medina, Tony</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 12-13</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Filly Willy Flim-flam Flyer Boost & Rocket Glider for Swift & SparroWevents, in Model Rocketeer Dec 1973 Volume 15 Number 11, Pages 12-13</a></li>
</ul>
<p><h1>Newill, David</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Model Rocket Tips (getting under-camber on wing), in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 13</a></li>
</ul>
<p><h1>Pearson, Ed</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">First World Model Rocket Championships-VRSAC '72, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 6-7, 14-15</a></li>
</ul>
<p><h1>Sadowski, Elaine</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">Editor's Nook, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 4</a></li>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Editor's Nook, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 5</a></li>
</ul>
<p><h1>Shenosky, Larry</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">Television and Model Rocketry, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 22</a></li>
</ul>
<p><h1>Unknown</h1></p>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR In Action, in Model Rocketeer Jan 1973 Volume 14 Number 1, Page 5</a></li>
<li><a href="http://fake.path/1973/MR-Jan1973_V14-N1.pdf">NAR News, in Model Rocketeer Jan 1973 Volume 14 Number 1, Pages 11, 14</a></li>
</ul>
<p><h1>Wright, Ron</h1></p>
<li><a href="http://fake.path/1973/MR-Dec1973_V15-N11.pdf">NAR In Action, in Model Rocketeer Dec 1973 Volume 15 Number 11, Page 20</a></li>
</ul>
"""

def count_rows(worksheet):
  row = 1
  while (worksheet['A' + str(row)].value != None):
    row += 1
  return row - 1

def convert_colnum_to_char(column):
  return string.ascii_uppercase[column - 1]

def count_cols(worksheet):
  col = 1
  while (worksheet[convert_colnum_to_char(col) + '1'].value != None):
    col += 1
  return col - 1

def get_row_col(worksheet, row, col):
  return worksheet[convert_colnum_to_char(col) + str(row)].value

def make_dict(worksheet):
  rows = count_rows(worksheet)
  cols = count_cols(worksheet)
  headers = {}
  headersToValues = {}
  for column in xrange(1, cols+1):
    header = get_row_col(worksheet, 1, column)
    headers[column] = header
    headersToValues[header] = {}
  for row in xrange(1, rows + 1):
    index = get_row_col(worksheet, row, 1)
    for col in headers.keys():
      if headers[col] == pageString:
        headersToValues[headers[col]][index] = \
            str(get_row_col(worksheet, row, col)).encode(
            'ascii','replace').replace('.0', '').replace('.', '')
      else:
        headersToValues[headers[col]][index] = get_row_col(worksheet, row, col)
  return headers, headersToValues

def make_issue_string(source, month, year, volume, issue):
  return (source + '-' + month + str(year) + '_V' + str(volume) + '-N' + \
      str(issue)).encode('ascii','replace')

def gather_issues(headers, headersToValues):
  issuesToIndices = collections.defaultdict(list)
  for index in headersToValues[indexString]:
    if index != u'IDX':
      issuesToIndices[make_issue_string(headersToValues[sourceString][index],
          headersToValues[monthString][index],
          headersToValues[yearString][index],
          headersToValues[volumeString][index],
          headersToValues[numberString][index])].append(index)
  return issuesToIndices
 
def break_issue_into_year_volume_number(issue):
  year = issue[6:10]
  volume = issue[string.rfind(issue, 'V') + 1:][:2]
  number = issue[string.rfind(issue, 'N') + 1:]
  return int(year), int(volume), int(number)
 
def sort_issues(issue1, issue2):
  year1, volume1, number1 = break_issue_into_year_volume_number(issue1)
  year2, volume2, number2 = break_issue_into_year_volume_number(issue2)
  if year1 > year2:
    return 1
  elif year2 > year1:
    return -1
  elif volume1 > volume2:
    return 1 
  elif volume2 > volume1:
    return -1   
  elif number1 > number2:
    return 1
  elif number1 == number2:
    return 0
  else:
    return -1    

def sort_all_issues(issues):
  issues.sort(sort_issues)
  return issues

def sort_indices_chronologically(indices, headerToValues):
  indicesYearVolumeNumberFirstPage = []
  for index in indices:
    indicesYearVolumeNumberFirstPage.append((index,
        headerToValues[yearString][index],
        headerToValues[volumeString][index],
        headerToValues[numberString][index],
        first_page(headerToValues[pageString][index])))
  indicesYearVolumeNumberFirstPage.sort(key = operator.itemgetter(1, 2, 3, 4))
  returnIndices = []
  for index, year, volume, number, page in indicesYearVolumeNumberFirstPage:
    returnIndices.append(index)
  return returnIndices

def issues_to_year_to_issues(issues):
  yearToIssues = collections.defaultdict(list)
  for issue in issues:
    year, volume, number = break_issue_into_year_volume_number(issue)
    yearToIssues[year].append(issue)   
  return yearToIssues

def pretty_name(issue):
  magName = issue[:string.find(issue, '-')]
  year, volume, number = break_issue_into_year_volume_number(issue)
  month = issue[3:6]
  return nameDict[magName] + ' ' + month + ' ' + str(year) + ' Volume ' + \
      str(volume) + ' Number ' + str(number)

def output_issue_html(issues, globalPath):
  yearsToIssues = issues_to_year_to_issues(issues)
  sortedYears = yearsToIssues.keys()
  sortedYears.sort()
  htmlString = ''
  for year in sortedYears:
    htmlString += '<p><h1>' + str(year) + '</h1></p>\n'
    for issue in sort_all_issues(yearsToIssues[year]):
      htmlString += '<p><a href="' + globalPath + str(year) + '/' + \
          issue + '.pdf">' + pretty_name(issue) + '</a></p>\n'
  return htmlString

def first_page(pageString):
  return int(pageString.split('-')[0].split(',')[0])

def unique_by_various(indices, headersToValues):
  outIndices = []
  for index in indices:
    add = True
    for out in outIndices:
      match = True
      for aString in [yearString, volumeString, numberString, titleString, authorString]:
        if headersToValues[aString][index] != headersToValues[aString][out]:
          match = False
      if match:
        add = False
    if add:
      outIndices.append(index)
  return outIndices

def sort_by_first_page(indices, headersToValues):
  indexPageTuples = []
  for index in indices:
    indexPageTuples.append(
        (index, first_page(headersToValues[pageString][index])))
  indexPageTuples.sort(key=operator.itemgetter(1))
  returnIndices = []
  for index, page in indexPageTuples:
    returnIndices.append(index)
  return returnIndices

def multiple_pages(page):
  return string.find(page, ',') != -1 or string.find(page, '-') != -1

def output_article_html(issues, issueToIndices, headersToValues, globalPath):
  yearsToIssues = issues_to_year_to_issues(issues)
  sortedYears = yearsToIssues.keys()
  sortedYears.sort()
  htmlString = ''
  for year in sortedYears:
    htmlString += '<p><h1>' + str(year) + '</h1></p>\n'
    for issue in sort_all_issues(yearsToIssues[year]):
      htmlString += '<p><h2><a href="' + globalPath + str(year) + '/' + \
          issue + '.pdf">' + pretty_name(issue) + '</a></h2></p>\n<ul>'
      for index in sort_by_first_page(issueToIndices[issue], headersToValues):
        htmlString += '<li><a href="' + globalPath + str(year) + '/' + \
            issue + '.pdf">' + \
            str(headersToValues[titleString][index]).encode('ascii','replace')+\
            ', by ' + \
            str(headersToValues[authorString][index]).encode('ascii','replace') + \
            ', Page'
        pageNum = str(headersToValues[pageString][index]).encode('ascii','replace')
        if multiple_pages(pageNum):
          htmlString += 's'
        htmlString +=  ' ' + pageNum + '</a></li>\n'
      htmlString += '</ul>\n'
  return htmlString

def output_article_html_unique(issues, issueToIndices, headersToValues, globalPath):
  yearsToIssues = issues_to_year_to_issues(issues)
  sortedYears = yearsToIssues.keys()
  sortedYears.sort()
  htmlString = ''
  for year in sortedYears:
    htmlString += '<p><h1>' + str(year) + '</h1></p>\n'
    for issue in sort_all_issues(yearsToIssues[year]):
      htmlString += '<p><h2><a href="' + globalPath + str(year) + '/' + \
          issue + '.pdf">' + pretty_name(issue) + '</a></h2></p>\n<ul>'
      firstPageSortedIndices = sort_by_first_page(issueToIndices[issue], headersToValues)
      for index in unique_by_various(firstPageSortedIndices, headersToValues):
        htmlString += '<li><a href="' + globalPath + str(year) + '/' + \
            issue + '.pdf">' + \
            str(headersToValues[titleString][index]).encode('ascii','replace')+\
            ', by ' + \
            str(headersToValues[authorString][index]).encode('ascii','replace') + \
            ', Page'
        pageNum = str(headersToValues[pageString][index]).encode('ascii','replace')
        if multiple_pages(pageNum):
          htmlString += 's'
        htmlString +=  ' ' + pageNum + '</a></li>\n'
      htmlString += '</ul>\n'
  return htmlString

def collect_indices_for_arbitrary(
    indices, arbitrary, headerToValues, arbitraryString):
  returnIndices = []
  for index in indices:
    if headerToValues[arbitraryString][index].encode(
        'ascii', 'replace') == arbitrary:
      returnIndices.append(index)
  return returnIndices

def collect_indices_for_category(indices, category, headerToValues):
  return collect_indices_for_arbitrary(
      indices, category, headerToValues, topicString)

def collect_indices_for_author(indices, author, headerToValues):
  return collect_indices_for_arbitrary(
      indices, author, headerToValues, authorString)

def find_issue(index, issueToIndices):
  for issue, indices in issueToIndices.iteritems():
    if index in indices:
      return issue
  return None

def output_category_html(issues, issueToIndices, headersToValues, globalPath):
  htmlString = ''
  indices = []
  for issue in issues:
    indices.extend(issueToIndices[issue])
  categories = set()
  for index in indices:
    categories.add(
        headersToValues[topicString][index].encode('ascii', 'replace'))
  categories = list(categories)
  categories.sort()
  for category in categories:
    htmlString += '<p><h1>' + category + '</h1></p>\n'
    indicesForCategory = collect_indices_for_category(
        indices, category, headersToValues)
    for index in sort_indices_chronologically(
        indicesForCategory, headersToValues):
      year = headersToValues[yearString][index]
      issue = find_issue(index, issueToIndices)
      htmlString += '<li><a href="' + globalPath + str(year) + '/' + \
          issue + '.pdf">' + \
          str(headersToValues[titleString][index]).encode('ascii','replace')+\
          ', by ' + \
          str(headersToValues[authorString][index]).encode('ascii','replace') + \
          ', in ' + pretty_name(issue) + \
          ', Page'
      pageNum = str(headersToValues[pageString][index]).encode('ascii','replace')
      if multiple_pages(pageNum):
        htmlString += 's'
      htmlString +=  ' ' + pageNum + '</a></li>\n'
    htmlString += '</ul>\n'
  return htmlString

def output_author_html(issues, issueToIndices, headersToValues, globalPath):
  htmlString = ''
  indices = []
  for issue in issues:
    indices.extend(issueToIndices[issue])
  authors = set()
  for index in indices:
    authors.add(
        headersToValues[authorString][index].encode('ascii', 'replace'))
  authors = list(authors)
  authors.sort()
  for author in authors:
    htmlString += '<p><h1>' + author + '</h1></p>\n'
    indicesForAuthor = collect_indices_for_author(
        indices, author, headersToValues)
    for index in sort_indices_chronologically(
        indicesForAuthor, headersToValues):
      year = headersToValues[yearString][index]
      issue = find_issue(index, issueToIndices)
      htmlString += '<li><a href="' + globalPath + str(year) + '/' + \
          issue + '.pdf">' + \
          str(headersToValues[titleString][index]).encode('ascii','replace')+\
          ', in ' + pretty_name(issue) + \
          ', Page'
      pageNum = str(headersToValues[pageString][index]).encode('ascii','replace')
      if multiple_pages(pageNum):
        htmlString += 's'
      htmlString +=  ' ' + pageNum + '</a></li>\n'
    htmlString += '</ul>\n'
  return htmlString

class TestProcessFunctions(unittest.TestCase):
  @classmethod
  def setUpClass(self):
    wb = openpyxl.load_workbook('test1973.xlsx', guess_types=True)
    self.ws = wb.active
    self.rowCount = count_rows(self.ws)
    self.colCount = count_cols(self.ws)
    self.headers, self.dict = make_dict(self.ws)  
    self.issuesToIndices = gather_issues(self.headers, self.dict)
    self.issueHtml = output_issue_html(self.issuesToIndices.keys(), testPath)
    outIssueHtml = open('chronological_by_issue.html', 'w')
    outIssueHtml.write(self.issueHtml)
    outIssueHtml.close()
    self.articleHtml = output_article_html(self.issuesToIndices.keys(), 
        self.issuesToIndices, self.dict, testPath)
    outArticleHtml = open('chronological_by_article.html', 'w')
    outArticleHtml.write(self.articleHtml)
    outArticleHtml.close()
    self.categoryHtml = output_category_html(self.issuesToIndices.keys(), 
        self.issuesToIndices, self.dict, testPath)
    outCategoryHtml = open('category.html', 'w')
    outCategoryHtml.write(self.categoryHtml)
    outCategoryHtml.close()
    self.authorHtml = output_author_html(self.issuesToIndices.keys(), 
        self.issuesToIndices, self.dict, testPath)
    outAuthorHtml = open('author.html', 'w')
    outAuthorHtml.write(self.authorHtml)
    outAuthorHtml.close()

  def test_count_rows(self):
    self.assertEqual(129, self.rowCount)
    
  def test_colnum_char_converter(self):
    self.assertEqual('A', convert_colnum_to_char(1))
    self.assertEqual('M', convert_colnum_to_char(13))
    self.assertEqual('Z', convert_colnum_to_char(26))

  def test_count_cols(self):
    self.assertEqual(11, self.colCount)

  def test_get_row_col(self):
    self.assertEqual('IDX', get_row_col(self.ws, 1, 1))
    self.assertEqual('AUTHOR', get_row_col(self.ws, 1, 2))
    self.assertEqual('Title', get_row_col(self.ws, 1, 11))

  def test_make_dict(self):
    headerGold = {1: u'IDX', 2: u'AUTHOR', 3: u'SOURCE', 4: u'Vol.', 5: u'No.', 6: u'Page', 7: u'Month', 8: u'# Month', 9: u'Year', 10: u'Topic', 11: u'Title'}
    self.assertEqual(headerGold, self.headers)
    self.assertIn(u'Title', self.dict.keys())
    self.assertIn(u'SOURCE', self.dict.keys())
    self.assertIn(u'Page', self.dict.keys())

  def test_make_issue_string(self):
    self.assertEqual('MR-Dec1973_V14-N12', make_issue_string('MR', 'Dec', 
        '1973', '14', '12'))

  def test_gather_issues(self):
    self.assertIn('MR-Dec1973_V15-N11', self.issuesToIndices.keys());
    self.assertEqual([2066, 3311, 3868, 389, 5086, 5091], self.issuesToIndices['MR-Apr1973_V15-N3'])
    self.assertTrue(set(['MR-Apr1973_V15-N3', 'MR-Aug1973_V15-N7',
        'MR-Dec1973_V15-N11', 'MR-Feb1973_V15-N1', 'MR-Jan1973_V14-N1',
        'MR-Jul1973_V15-N6', 'MR-Jun1973_V15-N5', 'MR-Mar1973_V15-N2',
        'MR-May1973_V15-N4', 'MR-Nov1973_V15-N10', 'MR-Oct1973_V15-N9',
        'MR-Sep1973_V15-N8']) == set(self.issuesToIndices.keys()))

  def test_break_issue_into_year_volume_number(self):
    self.assertEqual((1973, 15, 1), 
        break_issue_into_year_volume_number('MR-Feb1973_V15-N1'))

  def test_sort_issues(self):
    self.assertTrue(-1 == sort_issues('MR-Nov1973_V15-N10', 'MR-Dec1973_V15-N11'))
    self.assertTrue(-1 == sort_issues('MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'))
    self.assertTrue(0 == sort_issues('MR-Jan1973_V14-N1', 'MR-Jan1973_V14-N1'))
    self.assertTrue(1 == sort_issues('MR-Apr1973_V15-N3', 'MR-Feb1973_V15-N1'))
    self.assertTrue(-1 == sort_issues('MR-Jan1973_V14-N12', 'MR-Dec1973_V15-N11'))

  def test_sort_all_issues(self):
    self.assertEqual(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'], 
        sort_all_issues(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11']))
    self.assertEqual(['MR-Jan1973_V14-N1', 'MR-Feb1973_V15-N1',
        'MR-Mar1973_V15-N2', 'MR-Apr1973_V15-N3', 'MR-May1973_V15-N4',
        'MR-Jun1973_V15-N5', 'MR-Jul1973_V15-N6', 'MR-Aug1973_V15-N7',
        'MR-Sep1973_V15-N8', 'MR-Oct1973_V15-N9', 'MR-Nov1973_V15-N10',
        'MR-Dec1973_V15-N11'], sort_all_issues(self.issuesToIndices.keys()))

  def test_issues_to_year_to_issues(self):
    self.assertEqual({1973: ['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11']}, 
        issues_to_year_to_issues(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11']))

  def test_output_issue_html(self):
    self.assertEqual(goldIssueString,
        output_issue_html(
            ['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'], testPath))

  def test_output_article_html(self):
    self.assertEqual(goldArticleString,
        output_article_html(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'],
        self.issuesToIndices, self.dict, testPath))

  def test_output_article_html(self):
    print output_article_html_unique(
        ['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'],
        self.issuesToIndices, self.dict, testPath)
    self.assertEqual(goldArticleUniqueString,
        output_article_html_unique(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'],
        self.issuesToIndices, self.dict, testPath))

  def test_sort_indices_chronologically(self):
    self.assertEqual([2062, 3057],
        sort_indices_chronologically([2062, 3057], self.dict)) 
    self.assertEqual([182, 5088],
        sort_indices_chronologically([5088, 182], self.dict))

  def test_output_category_html(self):
    self.assertEqual(goldCategoryString,
        output_category_html(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'],
        self.issuesToIndices, self.dict, testPath))

  def test_output_author_html(self):
    self.assertEqual(goldAuthorString,
        output_author_html(['MR-Jan1973_V14-N1', 'MR-Dec1973_V15-N11'],
        self.issuesToIndices, self.dict, testPath))

if __name__ == '__main__':
  unittest.main()

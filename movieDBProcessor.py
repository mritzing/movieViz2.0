import xlsxwriter
import os
import urllib.request
import base64
outfile = 'movieData.xlsx'


os.remove(outfile)
f = open("movieNightHTML_src.html", "r")
splitArr = f.read().split("</poster-art>")


workbook = xlsxwriter.Workbook(outfile)
worksheet = workbook.add_worksheet()



count = 1
for element in splitArr:
	hrefIndex = element.find("<a ng-href=\"")
	hrefStr = element[hrefIndex+12:]
	#breaks on first and last element
	if hrefIndex is not 1:
			directStr = hrefStr[:hrefStr.find("\"")]
			
			thumbnailIndex = hrefStr[hrefStr.find("\""):].find("thumbnail-src=\"")
			thumbnailStr = hrefStr[thumbnailIndex+27:]
			thumbnailStr = thumbnailStr[:thumbnailStr.find("\"")]
			#rename when downloading to title
				
			titleIndex = hrefStr.find("title=\"")
			titleStr = hrefStr[titleIndex+7:]
			titleStr = titleStr[:titleStr.find("\"")]
			
			
			if thumbnailStr.find(".jpg") is not -1:
				saveStr = base64.urlsafe_b64encode(titleStr.encode())
				saveStr = "images/" + str(saveStr) + ".jpg"
				print (saveStr)
				urllib.request.urlretrieve(thumbnailStr, saveStr)


			worksheet.write(count, 0, titleStr)
			worksheet.write(count, 1, directStr)
			worksheet.write(count, 2, thumbnailStr)
	count+=1
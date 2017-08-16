from django.shortcuts import render
from django.template import loader
from django.urls import reverse
from django.http import HttpResponse
import requests

def index(request):
	#template_name = 'ExcelApp/main.html'
	#template = loader.get_template('ExcelApp/main.html')
	return render(request, 'ExcelApp/index.html')
	#return HttpResponse(template.render(request))

def details(request):
	return render(request, 'ExcelApp/main.html')


def getReport(request):
	rid = request.POST['rid']
	
	if rid == '2':
		response = requests.get('http://ykr-dev-apex.devyork.ca/apexenv/bsmart_data/stp_ws/stp_pop_test/')
		return HttpResponse(response)
		#return render(request, 'ExcelApp/index.html')
	else:
		return render(request, 'ExcelApp/main.html')
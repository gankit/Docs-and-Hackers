# -*- coding: utf-8 -*-
import scrapy
from scrapy import Request

import httplib2
import os
import re

class LabsSpider(scrapy.Spider):
    name = 'labs'
    allowed_domains = ['npidb.org']
    start_urls = ['https://npidb.org/doctors/allopathic_osteopathic_physicians/anatomic-pathology-clinical-pathology_207zp0102x/']

    def parse_page(self, response):
    	absolute_url = response.meta.get('URL')
    	name = response.meta.get('Name')
    	city = response.meta.get('City')
    	state = response.meta.get('State')
    	phone = response.xpath('//span[@itemprop="telephone"]/text()').extract_first()
    	authorized_official = ''

    	rows = response.xpath('//tr')    	
    	print('Number of rows: ' + str(len(rows)))
    	for row in rows:
    		columns = row.xpath('td')
    		if 'Authorized official' in columns[0].xpath('string(.)').extract_first():
    			authorized_official = columns[1].xpath('string(.)').extract_first()
    			authorized_official = authorized_official.strip()
    			authorized_official = re.sub('[\t\n\r\f\v]+', '', authorized_official)
    	yield {'URL':absolute_url, 'Name':name, 'City':city, 'State':state, 'Authorized Official':authorized_official, 'Phone':phone}

    def parse(self, response):
    	print("===========")
    	tables = response.xpath('//table')
    	for table in tables:
    		rows = table.xpath('tbody/tr')
    		for row in rows:
    			columns = row.xpath('td')
    			location = columns[3].xpath('text()').extract_first()
    			city = location[:-4]
    			state = location[-2:]

    			span_class = columns[0].xpath('span/@class').extract_first()
    			is_organization = False
    			if 'glyphicon-briefcase' in span_class:
    				is_organization = True
    			name = columns[1].xpath('h2/a/text()').extract_first()
    			relative_url = columns[1].xpath('h2/a/@href').extract_first()
    			absolute_url = response.urljoin(relative_url)
    			specialty = columns[4].xpath('a/text()').extract_first()
    			is_path_lab = False
    			if '207ZP0102X' in specialty:
    				is_path_lab = True
    			if is_organization and is_path_lab:    				
    				yield Request(absolute_url, callback=self.parse_page, meta={'URL':absolute_url, 'Name':name, 'City':city, 'State':state})
		next_buttons = response.xpath('//ul[@class="pagination"]/li')
		last_button = next_buttons[len(next_buttons)-1]
		relative_next_url = last_button.xpath('a/@href').extract_first()
		absolute_next_url = response.urljoin(relative_next_url)
		print(absolute_next_url)

		yield Request(absolute_next_url, callback=self.parse)		
		return
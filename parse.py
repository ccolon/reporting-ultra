#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import csv
from datetime import datetime
import xlsxwriter

class SkipException(Exception):
    pass 

weekday_mapping = {
    '0': 'lundi',
    '1': 'mardi',
    '2': 'mercredi',
    '3': 'jeudi',
    '4': 'vendredi',
    '5': 'samedi',
    '6': 'dimanche'
}

month_mapping = {
    '1': 'janvier',
    '2': 'fevrier',
    '3': 'mars',
    '4': 'avril',
    '5': 'mai',
    '6': 'juin',
    '7': 'juillet',
    '8': 'aout',
    '9': 'septembre',
    '10': 'octobre',
    '11': 'novembre',
    '12': 'decembre'
}

def prepare_data(filename):

    def check_header_line(line):
        return 'Container' in line or 'Receipts from' in line

    def check_empty_line(line):
        return line == '\n'

    def check_useless_line(line):
        return 'BIN' in line

    def extract_date(line):
        t = re.search('([0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{2}:[0-9]{2}:[0-9]{2})\.[0-9]{3}', line)
        if t is None:
            raise SkipException('Error - ' + line) 
        return t.group(1)

    def extract_numero(line):
        t = re.search('\s\sNo\s([0-9]+)\s', line)
        if t is None:
            raise SkipException('Error - ' + line) 
        return int(t.group(1))

    def extract_count(line):
        t = re.search('\sCnt\s([0-9]+)\s', line)
        if t is None:
            raise SkipException('Error - ' + line) 
        return int(t.group(1))

    def extract_amount(line):
        t = re.search('\sAmount\s([0-9]+)\s', line)
        if t is None:
            raise SkipException('Error - ' + line) 
        return int(t.group(1))

    def extract_donation(line):
        t = re.search('BC\s\|\|', line)
        if t is not None:
            return True        
        return False

    line_errors = []
    line_filtereds = []
    line_parseds = []
    print filename
    with open(filename, 'r') as f:
        for line in f:
            if check_header_line(line):
                continue
            if check_empty_line(line):
                continue
            if check_useless_line(line):
                continue
            try:
                line_parsed = {
                    'date': extract_date(line),
                    'numero': extract_numero(line),
                    'count': extract_count(line),
                    'amount': extract_amount(line),
                    'donation': extract_donation(line)
                }
            except SkipException as e:
                line_errors.append(str(e))
                continue
            if not line_parsed['date'].startswith('2017'):
                line_filtereds.append(line_parsed)
                continue
            line_parseds.append(line_parsed)

    line_enricheds = []
    for line in line_parseds:
        line_enriched = dict(line)
        date = datetime.strptime(line_enriched['date'], "%Y-%m-%d_%H:%M:%S")
        line_enriched['datetime'] = date
        line_enriched['year'] = date.year
        line_enriched['month'] = date.month
        line_enriched['month_hr'] = month_mapping[str(date.month)]
        line_enriched['day'] = date.day
        line_enriched['hour'] = date.hour
        line_enriched['minute'] = date.minute
        line_enriched['weekday'] = date.weekday()
        line_enriched['weekday_hr'] = weekday_mapping[str(date.weekday())]        
        line_enricheds.append(line_enriched)

    return line_enricheds

def filter_dates(data, start_date, end_date):    
    data_filtered = []
    for item in data:        
        if item['datetime'] >= start_date and item['datetime'] <= end_date:
            data_filtered.append(item)
    return data_filtered

def filter_trials(data):
    # remove tickets with amount of 0
    data_filtered = []
    for item in data:
        if item['amount'] != 0:
            data_filtered.append(item)
    return data_filtered

# Aggregations

def nb_bottles_per_hour(data):
    result = {}
    for item in data:
        try:
            result[item['hour']] += item['count']
        except KeyError:
            result[item['hour']] = item['count']
    return result


def nb_bottles_per_weekday(data):
    result = {}
    for item in data:
        try:
            result[item['weekday']] += item['count']
        except KeyError:
            result[item['weekday']] = item['count']
    return result


def nb_bottles_per_month(data):
    result = {}
    for item in data:
        try:
            result[item['month']] += item['count']
        except KeyError as e:
            result[item['month']] = item['count']
    return result

def amount_per_month(data):
    result = {}
    for item in data:
        try:
            result[item['month']] += item['amount']
        except KeyError as e:
            result[item['month']] = item['amount']
    return result


def nb_bottles_total(data):
    result = 0
    for item in data:
        result += item['count']
    return result


def amount_total(data):
    result = 0
    for item in data:
        result += item['amount']
    return result / 100.0


def count_total(data):
    return len(data)


def average_amount_per_month(data):
    result = {}
    result_temp = {}
    for item in data:
        try:
            result_temp[item['month']].append(item)
        except KeyError as e:
            result_temp[item['month']] = [item]
    
    for month, items in result_temp.iteritems():
        result[month] = int(round(sum(item['amount'] for item in items) / float(len(items))))

    return result


def count_nb_tickets_per_amount(data):
    result = {}
    result_temp = {}
    for item in data:
        try:
            result_temp[item['amount']].append(item)
        except KeyError as e:
            result_temp[item['amount']] = [item]
    
    for amount, items in result_temp.iteritems():
        result[amount] = len(items)

    return result

def count_nb_tickets_donation(data):
        
    count_donation = 0
    count_coupon = 0
    amount_donation = 0
    amount_coupon = 0
    for item in data:
        if item['donation']:
            count_donation += 1
            amount_donation += item['amount']
        else:
            count_coupon += 1
            amount_coupon += item['amount']
    
    return {
        'Nombre de dons': count_donation,
        'Nombre de bons': count_coupon,
        'Montant des dons': amount_donation / 100.0,
        'Montant des bons': amount_coupon / 100.0,
        'Pourcentage en valeur %': "%.2f" % ((amount_donation * 1.0) / (amount_donation + amount_coupon) * 100),
        'Pourcentage en volume %': "%.2f" % ((count_donation * 1.0) / (count_donation + count_coupon) * 100)
    }

def count_nb_tickets_per_range_amount(data):
    result = {
        '1-20 ct': 0,
        '21-40 ct': 0,
        '41-60 ct': 0,
        '61-80 ct': 0,
        '81-100 ct': 0
    }
    result_temp = count_nb_tickets_per_amount(data)
    for k in result_temp.keys():
        if 0 < int(k) <= 20:
            result['1-20 ct'] += result_temp[k]
        elif 20 < int(k) <= 40:
            result['21-40 ct'] += result_temp[k]
        elif 40 < int(k) <= 60:
            result['41-60 ct'] += result_temp[k]
        elif 60 < int(k) <= 80:
            result['61-80 ct'] += result_temp[k]
        elif 80 < int(k) <= 100:
            result['81-100 ct'] += result_temp[k]

    return result



# OUTPUT XLS FILE


def write_xls(line_enricheds, filename):

    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Write Data Headers
    def write_data(worksheet, data):
        data_headers = ['date', 'numero', 'count', 'amount', 'year', 'month', 'day', 'hour', 'minute', 'weekday']
        row = 0
        col = 0
        for header in data_headers:
            worksheet.write(row, col, header)
            col += 1

        # Write Data
        row = 1
        col = 0

        # Iterate over the data and write it out row by row.
        for line in data:
            worksheet.write(row, col,     line['date'])
            worksheet.write(row, col + 1, line['numero'])
            worksheet.write(row, col + 2, line['count'])
            worksheet.write(row, col + 3, line['amount'])
            worksheet.write(row, col + 4, line['year'])
            #worksheet.write(row, col + 5, line['month'])
            worksheet.write(row, col + 5, line['month_hr'].encode('utf-8'))
            worksheet.write(row, col + 6, line['day'])
            worksheet.write(row, col + 7, line['hour'])
            worksheet.write(row, col + 8, line['minute'])
            #worksheet.write(row, col + 1, line['weekday'])
            worksheet.write(row, col + 9, line['weekday_hr'])
            row += 1

    write_data(worksheet, line_enricheds)

    # Write Data Headers
    def write_aggreg_nb_bottles_per_hour(worksheet, data):
        header = 'Nb bouteilles/heure'
        row = 0
        col = 11
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 11

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, line)
            worksheet.write(row, col + 1, data[line])
            row += 1

    write_aggreg_nb_bottles_per_hour(worksheet, nb_bottles_per_hour(line_enricheds))


    # Write Data Headers
    def write_aggreg_nb_bottles_per_weekday(worksheet, data):
        header = 'Nb bouteille/jour semaine '
        row = 0
        col = 14
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 14

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, weekday_mapping[str(line)])
            worksheet.write(row, col + 1, data[line])
            row += 1

    write_aggreg_nb_bottles_per_weekday(worksheet, nb_bottles_per_weekday(line_enricheds))


    # Write Data Headers
    def write_aggreg_nb_bottles_per_month(worksheet, data):
        header = 'Nb bouteille/mois'
        row = 0
        col = 17
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 17

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, month_mapping[str(line)])
            worksheet.write(row, col + 1, data[line])
            row += 1

    write_aggreg_nb_bottles_per_month(worksheet, nb_bottles_per_month(line_enricheds))


    def write_aggreg_amount_per_month(worksheet, data):
        header = 'Montant/mois'
        row = 0
        col = 20
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 20

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, month_mapping[str(line)])
            worksheet.write(row, col + 1, data[line] / 100.0)
            row += 1

    write_aggreg_amount_per_month(worksheet, amount_per_month(line_enricheds))


    # Write Data Headers
    def write_aggreg_total_nb_bottle(worksheet, data):
        header = 'Nb bouteilles total'
        row = 0
        col = 23
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 23

        worksheet.write(row, col, data)

    write_aggreg_total_nb_bottle(worksheet, nb_bottles_total(line_enricheds))

    # Write Data Headers
    def write_aggreg_total_amount(worksheet, data):
        header = 'Montant total Bons'
        row = 3
        col = 23
        worksheet.write(row, col, header)    

        # Write Data
        row = 4
        col = 23

        worksheet.write(row, col, data)

    write_aggreg_total_amount(worksheet, amount_total(line_enricheds))

    # Write Data Headers
    def write_aggreg_total_count(worksheet, data):
        header = 'Nb total bons emis'
        row = 6
        col = 23
        worksheet.write(row, col, header)    

        # Write Data
        row = 7
        col = 23

        worksheet.write(row, col, data)

    write_aggreg_total_count(worksheet, count_total(line_enricheds))

    # Write Data Headers
    def write_aggreg_average_amount_tickets_per_month(worksheet, data):
        
        header = 'Moyenne des bons/mois'
        row = 9
        col = 23
        worksheet.write(row, col, header)    

        # Write Data
        row = 10
        col = 23

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, month_mapping[str(line)])
            worksheet.write(row, col + 1, data[line])
            row += 1
      
    write_aggreg_average_amount_tickets_per_month(worksheet, average_amount_per_month(line_enricheds))


    # Write Data Headers
    def write_aggreg_nb_tickets_per_range_amount(worksheet, data):
        
        header = 'Montant ct / Nb Bons'
        row = 0
        col = 26
        worksheet.write(row, col, header)    

        # Write Data
        row = 1
        col = 26

        # Iterate over the data and write it out row by row.
        for line in sorted(data):
            
            worksheet.write(row, col, line)
            worksheet.write(row, col + 1, data[line])
            row += 1
      
    write_aggreg_nb_tickets_per_range_amount(worksheet, count_nb_tickets_per_range_amount(line_enricheds))


    def write_ratio_donation_over_total(worksheet, data):

        header = 'Ratio Don vs Total'
        row = 0
        col = 29
        worksheet.write(row, col, header)
        
        # Write Data
        row = 1
        col = 29
        # Iterate over the data and write it out row by row.
        for key, value in data.iteritems():            
            worksheet.write(row, col, key)
            worksheet.write(row, col + 1, value)
            row += 1

    write_ratio_donation_over_total(worksheet, count_nb_tickets_donation(line_enricheds))

    workbook.close()

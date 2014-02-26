# encoding: utf-8

require_relative '../lib/xlsx_inspector.rb'

myTest = XLSXInspector::Inspector.new()
puts myTest.inspect('200krows.xlsx')
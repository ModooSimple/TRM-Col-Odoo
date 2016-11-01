# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (C) 2016 Christian Camilo Camargo.
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as
#    published by the Free Software Foundation, either version 3 of the
#    License, or (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################
"""
Conventions:
    o_var: var type object
    s_var: var type string (char)
    i_var: var type integer
    f_var: var type float
Description:
    This module performs the upgrade of the exchange rate automatically according
    to the period defined on the cronjob set in the XML.
"""
from openerp import models, fields, api
from suds.client import Client
from datetime import datetime
import suds
import logging
import xml.etree.ElementTree as ET
_logger = logging.getLogger(__name__)
class trmColombian(models.Model):
    _inherit = "res.currency.rate"
    def _get_soap_trm(self):
        s_url = "http://obiee.banrep.gov.co/analytics/saw.dll?wsdl"
        o_client = Client(s_url, service = "SAWSessionService")
        s_session_id = o_client.service.logon("publico", "publico")
        o_client.set_options(service = "XmlViewService")
        o_report = {
            "reportPath": "/shared/Consulta Series Estadisticas desde Excel/1. Tasa de Cambio Peso Colombiano/1.1 TRM - Disponible desde el 27 de noviembre de 1991/1.1.3 Serie historica para un rango de fechas dado",
            "reportXml": "null"
        }
        o_options = {
            "async" : "false",
            "maxRowsPerPage" : "100",
            "refresh" : "true",
            "presentationInfo" : "true"
        }
        try:
            o_result_query = o_client.service.executeXMLQuery(o_report, "SAWRowsetData", o_options, s_session_id)
            o_client.set_options(service = "SAWSessionService")
            o_client.service.logoff(s_session_id)
            o_xml_data = ET.fromstring(o_result_query.rowset)
            return o_xml_data[0][1].text, float(o_xml_data[0][2].text)
        except suds.WebFault as detail:
            o_client.set_options(service = "SAWSessionService")
            o_client.service.logoff(s_session_id)
            _logger.critical("Error while working with BancoRep API: " + detail)
            return "", 0.0
    @api.model
    def get_colombian_trm(self):
        s_name_rate, f_trm = self._get_soap_trm()
        i_currency_id = self.env["res.currency"].search([("name", "in", ("COP", "cop"))])[0].id
        try:
            s_name_last_rate = self.search([("currency_id", "=", i_currency_id)], limit = 1, order = "name desc")[0].name
        except:
            s_name_last_rate = ""
        if f_trm > 0 and s_name_rate != s_name_last_rate[0: -9]:
            o_vals = {
                "rate": f_trm,
                "name": datetime.strptime(s_name_rate, "%Y-%m-%d"),
                "currency_id": i_currency_id
            }
            self.create(o_vals)
            _logger.info("New exchange rate created to date: " + s_name_rate + ", with value: " + str(f_trm))

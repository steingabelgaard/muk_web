###################################################################################
# 
#    Copyright (C) 2017 MuK IT GmbH
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
###################################################################################

import uuid
import logging
import mimetypes

import werkzeug

from odoo import _, http
from odoo.http import request

import urllib.request, urllib.parse, urllib.error
import urllib.request, urllib.error, urllib.parse


_logger = logging.getLogger(__name__)

try:
    import requests
except ImportError:
    _logger.warn('Cannot `import requests`.')

MIMETPYES = [
    'application/msword', 'application/ms-word', 'application/vnd.ms-word.document.macroEnabled.12',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.mspowerpoint',
    'application/vnd.ms-powerpoint', 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'application/vnd.ms-powerpoint.presentation.macroEnabled.12'
]



class MSOfficeParserController(http.Controller):
    
    @http.route('/web/preview/msoffice', auth="user", type='http')
    def preview_msoffice(self, url, **kw):    
        _logger.info('INPUT:  %s', url)
        if not bool(urllib.parse.urlparse(url).netloc):
            method, params = self._get_route(url)
            response = method(**params)
            if not response.status_code == 200:
                return self._make_error_response(response.status_code,response.description if hasattr(response, 'description') else _("Unknown Error"))
            else:
                content_type = response.headers['content-type']
                data = response.data
        else:
            try:
                response = requests.get(url)
                content_type = response.headers['content-type']
                data = response.content
            except requests.exceptions.RequestException as exception:
                return self._make_error_response(exception.response.status_code, exception.response.reason or _("Unknown Error"))

        try:
            url = request.env['ir.config_parameter'].sudo().get_param('muk_web_preview.msoffice.pdf', 'http://converter/unoconv/pdf')
            files = {'file': data}
            _logger.info('REQUEST: %s', url)
            r = requests.post(url, files=files)
            filename = "%s%s" % (uuid.uuid4(), mimetypes.guess_extension(content_type))
            output = r.content
            _logger.info('OUTPUT: %d, %s', len(output), filename)
            return self._make_pdf_response(output, "%s.pdf" % filename)
        except Exception:
            _logger.exception("Error while convert the file.")
            return werkzeug.exceptions.InternalServerError()
    
    def _make_pdf_response(self, file, filename):
        headers = [('Content-Type', 'application/pdf'),
                   ('Content-Disposition', 'attachment; filename="{}";'.format(filename)),
                   ('Content-Length', len(file))]
        return request.make_response(file, headers)
    
    def _get_route(self, url):
        url_parts = url.split('?')
        path = url_parts[0]
        query_string = url_parts[1] if len(url_parts) > 1 else None
        router = request.httprequest.app.get_db_router(request.db).bind('')
        match = router.match(path, query_args=query_string)
        method = router.match(path, query_args=query_string)[0]
        params = dict(urllib.parse.parse_qsl(query_string))
        if len(match) > 1:
            params.update(match[1])
        return method, params

    def _make_error_response(self, status, message):
        exception = werkzeug.exceptions.HTTPException()
        exception.code = status
        exception.description = message
        return exception


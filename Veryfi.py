from datetime import datetime

import veryfi
from veryfi.errors import BadRequest, ResourceNotFound, UnauthorizedAccessToken, UnexpectedHTTPMethod, AccessLimitReached, InternalError


class Veryfi:

    def __init__(self, elcor_invoice_manager):
        self.elcor_invoice_manager = elcor_invoice_manager 
        pass

    def setup_veryfi(self, client_id, client_secret, username, api_key):
        return veryfi.Client(client_id, client_secret, username, api_key)
    
    
    def parse_veryfi(self, page, client):
        try:
            results = client.process_document(page)
        except ConnectionError:
            return 408
        except UnauthorizedAccessToken:
            return 401
        except BadRequest or ResourceNotFound or UnexpectedHTTPMethod or AccessLimitReached or InternalError:
            return 500

        try:
            unformatted_date = results['date']
            date = datetime.strptime(unformatted_date, '%Y-%m-%d %H:%M:%S')
            company = results['vendor']['name']
            items = results['line_items']
            concepts = []
            total = float(results['total']) * (-1)
        except:
            return None
        
        for item in items:
            item_desc = item.get('description')
            concepts.append(item_desc) if len(item_desc) < 25 else concepts.append(item_desc[:25])
            
        concepts = '/'.join(concepts)

        info = {'date': date, 'company': company, 'concepts': concepts, 'total': total}

        return info
    
    
    def check_response(self, filename, response):
        if isinstance(response, int) or response == None:
            self.elcor_invoice_manager.console.write(f'Could not parse {filename}\n')
            if response == 401:
                self.elcor_invoice_manager.console.write('Check your credential keys and client status at https://app.veryfi.com/\n')
            elif response == 408:
                self.elcor_invoice_manager.console.write('Check your internet connection.\n')
            elif response == None:
                self.elcor_invoice_manager.console.write('Provide a clearer document and try again.\n')
            else:
                self.elcor_invoice_manager.console.write('Something went wrong. Please try again later.\n')

            return False
        
        return True
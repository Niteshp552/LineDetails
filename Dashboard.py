import csv
from io import StringIO
from flask import Flask, request, render_template, make_response
from googleads import ad_manager
import traceback

app = Flask(__name__)


def initialize_ad_manager_client():
    """Initializes the Ad Manager client using the configuration in googleads.yaml."""
    try:
        client = ad_manager.AdManagerClient.LoadFromStorage(
            '/Users/Nitesh.Pandey1/Desktop/ColombiaQC/pythonProject/google-ads.yaml')
        return client
    except Exception as e:
        print("Failed to initialize Ad Manager client:")
        print(traceback.format_exc())
        raise


def fetch_order_details(client, order_id):
    """Fetches details of a specific order using the Ad Manager API."""
    try:
        order_service = client.GetService('OrderService', version='v202305')
        statement = (ad_manager.StatementBuilder()
                     .Where('id = :id')
                     .WithBindVariable('id', order_id)
                     .Limit(1))
        response = order_service.getOrdersByStatement(statement.ToStatement())
        if 'results' in response and len(response['results']):
            return response['results'][0]
        else:
            return None
    except Exception as e:
        print("Failed to fetch order details:")
        print(traceback.format_exc())
        raise


def fetch_line_items_for_order(client, order_id):
    """Fetches line items for a specific order using the Ad Manager API."""
    try:
        line_item_service = client.GetService('LineItemService', version='v202305')
        statement = (ad_manager.StatementBuilder()
                     .Where('orderId = :orderId')
                     .WithBindVariable('orderId', order_id))
        response = line_item_service.getLineItemsByStatement(statement.ToStatement())
        if 'results' in response and len(response['results']):
            return response['results']
        else:
            return None
    except Exception as e:
        print("Failed to fetch line items:")
        print(traceback.format_exc())
        raise


def fetch_inventory_for_line_items(client, line_items):
    """Fetches inventory details for the given line items using the Ad Manager API."""
    try:
        inventory_service = client.GetService('InventoryService', version='v202305')
        inventory_details = {}
        for line_item in line_items:
            if hasattr(line_item.targeting, 'inventoryTargeting'):
                inventory_details[line_item.id] = []
                for targeted_ad_unit in line_item.targeting.inventoryTargeting.targetedAdUnits:
                    ad_unit_id = targeted_ad_unit.adUnitId
                    statement = (ad_manager.StatementBuilder()
                                 .Where('id = :id')
                                 .WithBindVariable('id', ad_unit_id))
                    response = inventory_service.getAdUnitsByStatement(statement.ToStatement())
                    if 'results' in response and len(response['results']):
                        ad_unit = response['results'][0]
                        inventory_details[line_item.id].append(ad_unit.name)
        return inventory_details
    except Exception as e:
        print("Failed to fetch inventory details:")
        print(traceback.format_exc())
        raise


def format_custom_targeting(custom_targeting):
    if custom_targeting is None:
        return "N/A"

    if hasattr(custom_targeting, 'children'):
        expressions = []
        for child in custom_targeting.children:
            if hasattr(child, 'keyId') and hasattr(child, 'valueIds'):
                expressions.append(f"Key {child.keyId}: Values {child.valueIds}")
        return "; ".join(expressions)

    return "N/A"


def format_frequency_cap(frequency_caps):
    if not frequency_caps:
        return "N/A"

    caps = []
    for cap in frequency_caps:
        caps.append(f"{cap.maxImpressions} impressions per {cap.numTimeUnits} {cap.timeUnit}")

    return "; ".join(caps)


def format_line_items(line_items, inventory_details):
    formatted_items = []
    for line_item in line_items:
        formatted_item = {
            'id': line_item.id,
            'name': line_item.name,
            'status': line_item.status,
            'start_date': f"{line_item.startDateTime.date.year}-{line_item.startDateTime.date.month}-{line_item.startDateTime.date.day}",
            'end_date': f"{line_item.endDateTime.date.year}-{line_item.endDateTime.date.month}-{line_item.endDateTime.date.day}",
            'budget': f"{line_item.budget.currencyCode} {line_item.budget.microAmount / 1000000}",
            'cost_type': line_item.costType,
            'rate': f"{line_item.costPerUnit.currencyCode} {line_item.costPerUnit.microAmount / 1000000}" if line_item.costPerUnit else "N/A",
            'creative_size': ', '.join([f"{cp.size.width}x{cp.size.height}" if cp.size else "N/A" for cp in
                                        line_item.creativePlaceholders]) if line_item.creativePlaceholders else "N/A",
            'country': ', '.join([loc.displayName for loc in line_item.targeting.geoTargeting.targetedLocations if
                                  loc.type == 'COUNTRY']) if line_item.targeting.geoTargeting and line_item.targeting.geoTargeting.targetedLocations else "N/A",
            'inventory': ', '.join(inventory_details[line_item.id]) if line_item.id in inventory_details else "N/A",
            'custom_targeting': format_custom_targeting(
                line_item.targeting.customTargeting) if line_item.targeting and hasattr(line_item.targeting,
                                                                                        'customTargeting') else "N/A",
            'frequency_cap': format_frequency_cap(line_item.frequencyCaps) if line_item.frequencyCaps else "N/A"
        }
        formatted_items.append(formatted_item)
    return formatted_items


@app.route('/')
def index():
    return render_template('Homepage.html')


@app.route('/fetch_order', methods=['POST'])
def fetch_order():
    try:
        order_id = request.form['order_id']
        client = initialize_ad_manager_client()
        order = fetch_order_details(client, order_id)
        if order:
            line_items = fetch_line_items_for_order(client, order_id)
            inventory_details = fetch_inventory_for_line_items(client, line_items)
            formatted_line_items = format_line_items(line_items, inventory_details)
            return render_template('Index.html', order=order, line_items=formatted_line_items)
        else:
            return "Order not found", 404
    except Exception as e:
        print("An error occurred while fetching order details:")
        print(traceback.format_exc())
        return "An error occurred", 500


@app.route('/download_csv', methods=['GET'])
def download_csv():
    try:
        order_id = request.args.get('order_id')
        client = initialize_ad_manager_client()
        order = fetch_order_details(client, order_id)
        if order:
            line_items = fetch_line_items_for_order(client, order_id)
            inventory_details = fetch_inventory_for_line_items(client, line_items)
            formatted_line_items = format_line_items(line_items, inventory_details)

            # Create a CSV file in memory
            si = StringIO()
            writer = csv.writer(si)

            # Write CSV headers
            writer.writerow(['Line Item ID', 'Name', 'Status', 'Start Date', 'End Date', 'Budget', 'Cost Type', 'Rate',
                             'Creative Size', 'Country', 'Inventory', 'Custom Targeting', 'Frequency Cap'])

            # Write CSV data
            for item in formatted_line_items:
                writer.writerow([
                    item['id'],
                    item['name'],
                    item['status'],
                    item['start_date'],
                    item['end_date'],
                    item['budget'],
                    item['cost_type'],
                    item['rate'],
                    item['creative_size'],
                    item['country'],
                    item['inventory'],
                    item['custom_targeting'],
                    item['frequency_cap'],
                ])

            # Serve the CSV file as a response
            response = make_response(si.getvalue())
            response.headers['Content-Disposition'] = 'attachment; filename=order_data.csv'
            response.headers["Content-type"] = "text/csv"
            return response
        else:
            return "Order not found", 404
    except Exception as e:
        print("An error occurred while generating CSV:")
        print(traceback.format_exc())
        return "An error occurred", 500


if __name__ == '__main__':
    app.run(debug=True)

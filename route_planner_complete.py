import time
import pandas as pd
import folium
from folium import plugins
from folium import PolyLine, Marker
from folium.plugins import MarkerCluster
import openrouteservice
from openrouteservice import convert
import webbrowser
from jinja2 import Template, Environment
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog
import os
import geopy.distance

# Read the Excel file
def read_excel(file_path):
    print("Reading Excel file...")
    df = pd.read_excel(file_path)
    df = df[df['Region'] == "BKK"]
    return df

# Sort orders into trucks based on weight
def sort_orders_into_trucks(df, max_weight=350):
    print("Sorting orders into trucks...")
    trucks = []
    current_truck = []
    current_weight = 0
    
    for index, row in df.iterrows():
        if current_weight + row['Total weights per order'] <= max_weight:
            current_truck.append(row)
            current_weight += row['Total weights per order']
        else:
            trucks.append(current_truck)
            current_truck = [row]
            current_weight = row['Total weights per order']
    
    if current_truck:
        trucks.append(current_truck)
    
    return trucks

# Function to calculate distance between two coordinates
def calculate_distance(coord1, coord2):
    return geopy.distance.distance(coord1, coord2).km

# Function to find the nearest order
def find_nearest_order(current_location, remaining_orders):
    nearest_order = None
    nearest_order_idx = None
    shortest_distance = float('inf')
    
    for idx, order in remaining_orders.iterrows():
        distance = calculate_distance(current_location, (order['Lat'], order['Lng']))
        if distance < shortest_distance:
            shortest_distance = distance
            nearest_order = order
            nearest_order_idx = idx
            
    return nearest_order, nearest_order_idx

# Function to reorder orders based on nearest neighbor approach
def reorder_orders(orders, source):
    print("Reordering orders...")
    reordered_orders = []
    current_location = source
    remaining_orders = orders.copy()
    
    while not remaining_orders.empty:
        nearest_order, nearest_order_idx = find_nearest_order(current_location, remaining_orders)
        reordered_orders.append(nearest_order)
        current_location = (nearest_order['Lat'], nearest_order['Lng'])
        remaining_orders = remaining_orders.drop(nearest_order_idx)
    
    return reordered_orders

# Optimize route using OpenRouteService
def optimize_route(client, source, orders, max_distance=6000000, radiuses=3000):
    reordered_orders = reorder_orders(orders, source)
    coordinates = [[source[1], source[0]]] + [[order['Lng'], order['Lat']] for order in reordered_orders]
    total_distance = 0
    
    for i in range(len(coordinates) - 1):
        total_distance += calculate_distance((coordinates[i][1], coordinates[i][0]), (coordinates[i+1][1], coordinates[i+1][0]))
        if total_distance > max_distance / 1000:  # Convert meters to kilometers
            split_index = i + 1
            break
    else:
        split_index = len(coordinates)
    
    chunks = [coordinates[i:i + split_index] for i in range(0, len(coordinates), split_index)]
    
    optimized_routes = []
    for chunk in chunks:
        try:
            optimized_route = client.directions(coordinates=chunk, profile='driving-hgv', format='geojson', radiuses=radiuses)
            optimized_routes.append(optimized_route)
        except openrouteservice.exceptions.ApiError as e:
            if 'rate limit' in str(e).lower():
                print("Rate limit exceeded, waiting for 60 seconds before retrying...")
                time.sleep(60)
                return optimize_route(client, source, orders, max_distance, radiuses)
            elif '2010' in str(e):  # Handle specific error for routable point
                if radiuses <= 3:
                    print(f"Radius {radiuses} too low, skipping optimization for this chunk.")
                    continue
                print(f"Error with radius {radiuses}, reducing radius and retrying...")
                return optimize_route(client, source, orders, max_distance, radiuses // 10)
            else:
                print(f"API error: {e}")
                raise e

    return optimized_routes, reordered_orders

# Calculate estimated time of arrival
def calculate_eta(optimized_routes, start_time):
    print("Calculating ETA...")
    eta = []
    current_time = start_time
    for optimized_route in optimized_routes:
        for segment in optimized_route['features'][0]['properties']['segments']:
            current_time += timedelta(seconds=segment['duration'])
            eta.append(current_time.strftime('%H:%M'))
    return eta

# Create map with polylines
def create_map(route, eta, map_obj, color, orders, truck_id):
    print("Creating map...")
    folium.Marker(location=[route[0][1], route[0][0]], popup="Start", icon=folium.Icon(color=color)).add_to(map_obj)
    
    for i, order in enumerate(orders):
        popup_text = f"Order: {order['Order Number']}<br>Destination: {order['Geo District']}<br>Weight: {order['Total weights per order']}<br>ETA: {eta[i]}"
        folium.Marker(location=[order['Lat'], order['Lng']], popup=popup_text, icon=plugins.BeautifyIcon(
                            icon="arrow-down", icon_shape="marker",
                            number=str(i+1),
                            background_color=color)).add_to(map_obj)
    
    polyline = PolyLine(locations=[[point[1], point[0]] for point in route], color=color)
    polyline.add_to(map_obj)

# Create sidebar with order details
def create_sidebar(trucks, all_eta, colors):
    print("Creating sidebar...")
    sidebar_html = """
    <div id="sidebar" style="position: fixed; top: 10px; left: 10px; width: 350px; height: 90%; overflow: auto; background: white; padding: 10px; border: 1px solid #ccc; z-index:9999;">
        <h3>Order Details</h3>
        {% for truck_idx, truck in trucks %}
            <h4><span style="display:inline-block; width:20px; height:10px; background-color:{{ colors[truck_idx % colors|length] }};"></span> Truck {{ truck_idx + 1 }}</h4>
            <p>Total Weight: {{ truck | sum(attribute='Total weights per order') }} kg</p>
            <p>Total Orders: {{ truck | length }}</p>
            <table border="1" style="width: 100%; margin-bottom: 20px;">
                <tr>
                    <th>Order</th>
                    <th>Destination</th>
                    <th>Weight</th>
                    <th>ETA</th>
                </tr>
                {% for order_idx, order in enumerate(truck) %}
                    <tr>
                        <td>{{ order['Order Number'] }}</td>
                        <td>{{ order['Geo District'] }}</td>
                        <td>{{ order['Total weights per order'] }}</td>
                        <td id="truck{{ truck_idx }}_eta_{{ order_idx }}">{{ all_eta[truck_idx][order_idx] }}</td>
                    </tr>
                {% endfor %}
            </table>
        {% endfor %}
    </div>
    """
    template = Environment().from_string(sidebar_html)
    return template.render(trucks=list(enumerate(trucks)), all_eta=all_eta, colors=colors, enumerate=enumerate)

# Main function to generate HTML map
def main_generate_html(file_path, ors_api_key):
    global trucks, all_eta, colors
    # Initialize OpenRouteService client
    client = openrouteservice.Client(key=ors_api_key)
    
    # Static source coordinates
    source = [14.0828151, 100.6258423]  # Note: lat, lng
    
    # Read Excel file
    df = read_excel(file_path)
    
    # Sort orders into trucks
    trucks = sort_orders_into_trucks(df)
    
    # Initialize maps
    main_map = folium.Map(location=[source[0], source[1]], zoom_start=12, control_scale=True)
    
    # Define colors for different trucks
    colors = ['blue', 'green', 'red', 'purple', 'orange', 'darkred', 'lightred', 'beige', 'darkblue', 'darkgreen', 'cadetblue', 'darkpurple', 'pink', 'lightblue', 'lightgreen', 'gray', 'black']
    
    all_eta = []
    start_time = datetime.strptime("11:00", "%H:%M")
    
    # Process each truck and create individual maps
    for i, truck in enumerate(trucks):
        optimized_routes, reordered_truck = optimize_route(client, source, pd.DataFrame(truck))
        eta = calculate_eta(optimized_routes, start_time)
        all_eta.append(eta)
        
        for optimized_route in optimized_routes:
            route = optimized_route['features'][0]['geometry']['coordinates']
            create_map(route, eta[:len(reordered_truck)], main_map, colors[i % len(colors)], reordered_truck, i)
    
    # Add sidebar with order details
    sidebar_html = create_sidebar(trucks, all_eta, colors)
    main_map.get_root().html.add_child(folium.Element(sidebar_html))
    
    # Save and open the main map
    main_map.save('route_map_all.html')
    webbrowser.open('route_map_all.html')
    
    # Enable truck selection dropdown
    dropdown_truck.configure(state="normal")
    dropdown_truck.configure(values=[f"Truck {i+1}" for i in range(len(trucks))])

# Function to generate HTML map for a specific truck
def generate_truck_map(selected_truck):
    if selected_truck:
        truck_idx = int(selected_truck.split()[1]) - 1
        truck = trucks[truck_idx]
        color = colors[truck_idx % len(colors)]
        
        # Static source coordinates
        source = [14.0828151, 100.6258423]  # Note: lat, lng
        
        truck_map = folium.Map(location=[source[0], source[1]], zoom_start=12, control_scale=True)
        
        optimized_routes, reordered_truck = optimize_route(openrouteservice.Client(key=entry_api_key.get()), source, pd.DataFrame(truck))
        eta = calculate_eta(optimized_routes, datetime.strptime("11:00", "%H:%M"))
        
        for optimized_route in optimized_routes:
            route = optimized_route['features'][0]['geometry']['coordinates']
            create_map(route, eta[:len(reordered_truck)], truck_map, color, reordered_truck, truck_idx)
        
        # Add sidebar with order details for specific truck
        sidebar_html = create_sidebar([reordered_truck], [eta], [color])
        truck_map.get_root().html.add_child(folium.Element(sidebar_html))
        
        print("Saving truck map...")
        truck_map.save(f'route_map_truck_{truck_idx+1}.html')
        webbrowser.open(f'route_map_truck_{truck_idx+1}.html')

# Export truck orders to Excel
def export_to_excel(file_path, output_path):
    df = read_excel(file_path)
    trucks = sort_orders_into_trucks(df)
    
    with pd.ExcelWriter(output_path) as writer:
        for i, truck in enumerate(trucks):
            truck_df = pd.DataFrame(truck)
            truck_df.to_excel(writer, sheet_name=f'Truck {i+1}', index=False)

# UI with customTkinter
def select_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        label_selected_file.configure(text=f"Selected File: {os.path.basename(file_path)}")

def generate_html():
    if file_path:
        ors_api_key = entry_api_key.get()
        main_generate_html(file_path, ors_api_key)

def export_excel():
    if file_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
        if output_path:
            export_to_excel(file_path, output_path)

# Initialize UI
app = ctk.CTk()
app.geometry("400x400")
app.title("Truck Route Planner")

file_path = ""

label_api_key = ctk.CTkLabel(app, text="OpenRouteService API Key:")
label_api_key.pack(pady=10)

entry_api_key = ctk.CTkEntry(app, width=300)
entry_api_key.pack(pady=5)

button_select_file = ctk.CTkButton(app, text="Select Excel File", command=select_file)
button_select_file.pack(pady=10)

label_selected_file = ctk.CTkLabel(app, text="No file selected")
label_selected_file.pack(pady=5)

button_generate_html = ctk.CTkButton(app, text="Generate HTML Map", command=generate_html)
button_generate_html.pack(pady=10)

button_export_excel = ctk.CTkButton(app, text="Export to Excel", command=export_excel)
button_export_excel.pack(pady=10)

dropdown_truck = ctk.CTkOptionMenu(app, values=[], state="disabled", command=generate_truck_map)
dropdown_truck.pack(pady=10)

app.mainloop()

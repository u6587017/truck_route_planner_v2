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
from tkinter import filedialog, messagebox
import os
import geopy.distance

# Read the Excel file
def read_excel(file_path):
    print("Reading Excel file...")
    df = pd.read_excel(file_path)
    df = df[df['Region'] == "BKK"]
    return df

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

# Sort orders into trucks based on nearest neighbor approach and weight limit
def sort_orders_into_trucks(df, max_weight=350):
    print("Sorting orders into trucks...")
    trucks = []
    current_truck = []
    current_weight = 0
    remaining_orders = df.copy()
    source = [14.0828151, 100.6258423]  # Note: lat, lng
    
    while not remaining_orders.empty:
        nearest_order, nearest_order_idx = find_nearest_order(source, remaining_orders)
        if current_weight + nearest_order['Total weights per order'] <= max_weight:
            current_truck.append(nearest_order)
            current_weight += nearest_order['Total weights per order']
        else:
            trucks.append(current_truck)
            current_truck = [nearest_order]
            current_weight = nearest_order['Total weights per order']
        
        source = [nearest_order['Lat'], nearest_order['Lng']]
        remaining_orders = remaining_orders.drop(nearest_order_idx)
    
    if current_truck:
        trucks.append(current_truck)
    
    return trucks

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
def optimize_route(client, source, orders, max_distance=6000000, radiuses=3000, profile='driving-hgv'):
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
            optimized_route = client.directions(coordinates=chunk, profile=profile, format='geojson', radiuses=radiases)
            optimized_routes.append(optimized_route)
        except openrouteservice.exceptions.ApiError as e:
            if 'rate limit' in str(e).lower():
                print("Rate limit exceeded, waiting for 60 seconds before retrying...")
                time.sleep(60)
                return optimize_route(client, source, orders, max_distance, radiases, profile)
            elif '2010' in str(e):  # Handle specific error for routable point
                if radiases <= 3:
                    print(f"Radius {radiases} too low, skipping optimization for this chunk.")
                    continue
                print(f"Error with radius {radiases}, reducing radius and retrying...")
                return optimize_route(client, source, orders, max_distance, radiases // 10, profile)
            elif '2009' in str(e):  # Handle specific error for route not found
                print(f"Route could not be found: {e}")
                continue
            else:
                print(f"API error: {e}")
                raise e

    return optimized_routes, reordered_orders

# Calculate estimated time of arrival
def calculate_eta(optimized_routes, start_time):
    print("Calculating ETA...")
    eta = []
    current_time = start_time
    max_eta_time = start_time.replace(hour=19, minute=0)
    
    for optimized_route in optimized_routes:
        for segment in optimized_route['features'][0]['properties']['segments']:
            if current_time > max_eta_time:
                return eta
            current_time += timedelta(seconds=segment['duration'])
            current_time += timedelta(seconds=1800)  # Add 30 minutes
            eta.append(current_time.strftime('%H:%M'))
    
    return eta

# Create map with polylines
def create_map(route, eta, map_obj, color, orders, truck_id):
    print("Creating map...")
    if route:
        folium.Marker(location=[route[0][1], route[0][0]], popup="Start", icon=folium.Icon(color=color)).add_to(map_obj)
    
    for i, order in enumerate(orders):
        popup_text = f"Order: {order['Order Number']}<br>Destination: {order['Geo District']}<br>Weight: {order['Total weights per order']}<br>ETA: {eta[i] if i < len(eta) else 'N/A'}"
        folium.Marker(location=[order['Lat'], order['Lng']], popup=popup_text, icon=plugins.BeautifyIcon(
                            icon="arrow-down", icon_shape="marker",
                            number=str(i+1),
                            background_color=color)).add_to(map_obj)
    
    if route:
        polyline = PolyLine(locations=[[point[1], point[0]] for point in route], color=color)
        polyline.add_to(map_obj)

# Create sidebar with order details
def create_sidebar(trucks, all_eta, colors):
    print("Creating sidebar...")
    sidebar_html = """
    <div id="sidebar" style="position: fixed; top: 10px; left: 10px; width: 350px; height: 90%; overflow: auto; background: white; padding: 10px; border: 1px solid #000; z-index:9999;">
        <h3>Order Details</h3>
        {% for truck_idx, truck in trucks %}
            <h4><span style="display:inline-block; width:20px; height:10px; background-color:{{ colors[truck_idx % colors|length] }};"></span> Truck {{ truck_idx + 1 }}</h4>
            <p>Total Weight: {{ truck | sum(attribute='Total weights per order') }} kg</p>
            <p>Total Orders: {{ truck | length }}</p>
            <table border="1" style="border-collapse: collapse; width: 100%; margin-bottom: 20px;">
                <tr>
                    <th style="border: 1px solid #000;">#</th>
                    <th style="border: 1px solid #000;">Order</th>
                    <th style="border: 1px solid #000;">Destination</th>
                    <th style="border: 1px solid #000;">Weight</th>
                    <th style="border: 1px solid #000;">ETA</th>
                </tr>
                {% for order_idx, order in enumerate(truck) %}
                    <tr>
                        <td style="border: 1px solid #000;">{{ order_idx + 1 }}</td>
                        <td style="border: 1px solid #000;">{{ order['Order Number'] }}</td>
                        <td style="border: 1px solid #000;">{{ order['Geo District'] }}</td>
                        <td style="border: 1px solid #000;">{{ order['Total weights per order'] }}</td>
                        <td style="border: 1px solid #000;" id="truck{{ truck_idx }}_eta_{{ order_idx }}">{{ all_eta[truck_idx][order_idx] if order_idx < len(all_eta[truck_idx]) else 'N/A' }}</td>
                    </tr>
                {% endfor %}
            </table>
        {% endfor %}
    </div>
    """
    env = Environment()
    env.globals.update({'len': len})
    template = env.from_string(sidebar_html)
    return template.render(trucks=list(enumerate(trucks)), all_eta=all_eta, colors=colors, enumerate=enumerate)

# Main function to generate HTML map
def main_generate_html(file_path, ors_api_key, start_time_str):
    global trucks, all_eta, colors, folium_colors
    # Check if the API key is provided
    if not ors_api_key:
        print("API key is missing. Please provide a valid OpenRouteService API key.")
        return
    
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
    
    # Define colors and corresponding color codes for different trucks
    colors = ['#0000FF', '#008000', '#FF0000', '#800080', '#FFA500', '#8B0000', '#FF6347', '#F5F5DC', '#00008B', '#006400', '#5F9EA0', '#4B0082', '#FFC0CB', '#ADD8E6', '#90EE90', '#808080', '#000000']
    folium_colors = {'#0000FF', '#008000', '#FF0000', '#800080', '#FFA500', '#8B0000', '#FF6347', '#F5F5DC', '#00008B', '#006400', '#5F9EA0', '#4B0082', '#FFC0CB', '#ADD8E6', '#90EE90', '#808080', '#000000'}
    
    all_eta = []
    
    # Convert start time string to datetime object
    start_time = datetime.strptime(start_time_str, "%H:%M")
    
    trucks_with_eta = []
    trucks_without_eta = []

    # Process each truck and create individual maps
    for i, truck in enumerate(trucks):
        optimized_routes, reordered_truck = optimize_route(client, source, pd.DataFrame(truck))
        if optimized_routes:
            eta = calculate_eta(optimized_routes, start_time)
            all_eta.append(eta)
            trucks_with_eta.append(reordered_truck)
        else:
            trucks_without_eta.extend(truck)
            all_eta.append(['N/A'] * len(truck))
        
        for optimized_route in optimized_routes:
            route = optimized_route['features'][0]['geometry']['coordinates']
            color = colors[i % len(colors)]
            color = color if color in folium_colors else '#0000FF'
            create_map(route, eta[:len(reordered_truck)], main_map, color, reordered_truck, i)
    
    # Reprocess trucks without ETA
    if trucks_without_eta:
        print("Reprocessing trucks without ETA...")
        trucks_without_eta_df = pd.DataFrame(trucks_without_eta)
        optimized_routes, reordered_truck = optimize_route(client, source, trucks_without_eta_df, profile='driving-car')
        if optimized_routes:
            eta = calculate_eta(optimized_routes, start_time)
            all_eta.append(eta)
            trucks_with_eta.append(reordered_truck)
        else:
            all_eta.append(['N/A'] * len(trucks_without_eta_df))

        for optimized_route in optimized_routes:
            route = optimized_route['features'][0]['geometry']['coordinates']
            color = colors[len(trucks_with_eta) % len(colors)]
            color = color if color in folium_colors else '#0000FF'
            create_map(route, eta[:len(reordered_truck)], main_map, color, reordered_truck, len(trucks_with_eta))
    
    # Add sidebar with order details
    sidebar_html = create_sidebar(trucks_with_eta, all_eta, colors)
    main_map.get_root().html.add_child(folium.Element(sidebar_html))
    
    # Save and open the main map
    main_map.save('route_map_all.html')
    webbrowser.open('route_map_all.html')
    
    # Enable truck selection dropdown
    dropdown_truck.configure(state="normal")
    dropdown_truck.configure(values=[f"Truck {i+1}" for i in range(len(trucks_with_eta))])

# Function to generate HTML map for a specific truck
def generate_truck_map(selected_truck):
    global color_codes, folium_colors
    if selected_truck:
        truck_idx = int(selected_truck.split()[1]) - 1
        truck = trucks[truck_idx]
        color = colors[truck_idx % len(colors)]
        color = color if color in folium_colors else '#0000FF'
        
        # Static source coordinates
        source = [14.0828151, 100.6258423]  # Note: lat, lng
        
        truck_map = folium.Map(location=[source[0], source[1]], zoom_start=12, control_scale=True)
        
        optimized_routes, reordered_truck = optimize_route(openrouteservice.Client(key=entry_api_key.get()), source, pd.DataFrame(truck))
        eta = calculate_eta(optimized_routes, datetime.strptime(entry_start_time.get(), "%H:%M")) if optimized_routes else ['N/A'] * len(truck)
        
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
    global trucks, all_eta
    trucks = []
    all_eta = []
    if file_path:
        ors_api_key = entry_api_key.get()
        start_time_str = entry_start_time.get()
        if ors_api_key and start_time_str:
            try:
                datetime.strptime(start_time_str, "%H:%M")
            except ValueError:
                messagebox.showerror("Invalid Time Format", "Please enter a valid start time in HH:MM format.")
                return
            main_generate_html(file_path, ors_api_key, start_time_str)
        else:
            print("Please provide a valid OpenRouteService API key and start time.")

def export_excel():
    if file_path:
        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")])
        if output_path:
            export_to_excel(file_path, output_path)

# Initialize UI
app = ctk.CTk()
app.geometry("400x500")
app.title("Truck Route Planner")

file_path = ""

label_api_key = ctk.CTkLabel(app, text="OpenRouteService API Key:")
label_api_key.pack(pady=10)

entry_api_key = ctk.CTkEntry(app, width=300)
entry_api_key.pack(pady=5)

label_start_time = ctk.CTkLabel(app, text="Start Time (HH:MM):")
label_start_time.pack(pady=10)

entry_start_time = ctk.CTkEntry(app, width=300)
entry_start_time.pack(pady=5)

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

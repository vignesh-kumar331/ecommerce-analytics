import csv
import random
from datetime import datetime, timedelta

random.seed(42)

NUM_CUSTOMERS = 2500
NUM_ORDERS = 6000
DATE_START = datetime(2024, 7, 1)
DATE_END = datetime(2025, 12, 31)

CATEGORIES = {
    "Electronics": {"min": 50, "max": 500, "weight": 0.28},
    "Clothing": {"min": 15, "max": 200, "weight": 0.25},
    "Home & Kitchen": {"min": 25, "max": 300, "weight": 0.22},
    "Beauty": {"min": 10, "max": 150, "weight": 0.15},
    "Sports": {"min": 20, "max": 250, "weight": 0.10},
}

PRODUCTS = {
    "Electronics": ["Wireless Earbuds", "Bluetooth Speaker", "USB-C Hub", "Webcam", "Portable Charger", "Laptop Stand", "Mechanical Keyboard", "Mouse Pad", "HDMI Cable", "Smart Watch Band"],
    "Clothing": ["Cotton T-Shirt", "Denim Jacket", "Running Shorts", "Casual Hoodie", "Formal Shirt", "Wool Scarf", "Cargo Pants", "Polo Shirt", "Track Pants", "Linen Shirt"],
    "Home & Kitchen": ["Non-Stick Pan Set", "Coffee Maker", "Air Purifier", "LED Desk Lamp", "Bamboo Cutting Board", "Insulated Flask", "Knife Set", "Spice Rack", "Storage Containers", "Blender"],
    "Beauty": ["Moisturizer", "Sunscreen SPF50", "Vitamin C Serum", "Lip Balm Set", "Face Wash", "Hair Oil", "Sheet Mask Pack", "Body Lotion", "Perfume Roll-On", "Eye Cream"],
    "Sports": ["Yoga Mat", "Resistance Bands", "Jump Rope", "Water Bottle", "Dumbbell Set", "Gym Bag", "Running Armband", "Foam Roller", "Swimming Goggles", "Sports Towel"],
}


def random_date(start, end):
    delta = end - start
    return start + timedelta(days=random.randint(0, delta.days))


def generate_acquisition_dates():
    dates = []
    for _ in range(NUM_CUSTOMERS):
        acquisition = random_date(DATE_START, DATE_END - timedelta(days=30))
        dates.append(acquisition)
    return dates


def pick_category():
    cats = list(CATEGORIES.keys())
    weights = [CATEGORIES[c]["weight"] for c in cats]
    return random.choices(cats, weights=weights, k=1)[0]


def generate_orders():
    acquisition_dates = generate_acquisition_dates()

    orders = []
    order_id = 1000

    for _ in range(NUM_ORDERS):
        cid = random.randint(1, NUM_CUSTOMERS)
        acquisition = acquisition_dates[cid - 1]

        order_date = random_date(DATE_START, DATE_END)

        category = pick_category()
        product = random.choice(PRODUCTS[category])
        price_range = CATEGORIES[category]
        amount = round(random.uniform(price_range["min"], price_range["max"]), 2)

        orders.append({
            "customer_id": f"C-{cid:04d}",
            "order_id": f"O-{order_id:04d}",
            "order_date": order_date.strftime("%m/%d/%Y"),
            "order_amount": amount,
            "product_category": category,
            "product_name": product,
            "customer_acquisition_date": acquisition.strftime("%m/%d/%Y"),
        })
        order_id += 1

    orders.sort(key=lambda x: x["order_date"])
    return orders


def main():
    orders = generate_orders()
    fieldnames = ["customer_id", "order_id", "order_date", "order_amount",
                  "product_category", "product_name", "customer_acquisition_date"]

    output_path = "orders.csv"
    with open(output_path, "w", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(orders)

    unique_customers = len(set(o["customer_id"] for o in orders))
    print(f"Generated {len(orders)} orders for {unique_customers} customers")
    print(f"Date range: {orders[0]['order_date']} to {orders[-1]['order_date']}")
    print(f"Saved to {output_path}")


if __name__ == "__main__":
    main()

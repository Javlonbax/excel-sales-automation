import pandas as pd
import random
from pathlib import Path


output_dir = Path("../data/input")
output_dir.mkdir(parents=True, exist_ok=True)


products = {
    101: ("Laptop", "Electronics"),
    102: ("Monitor", "Electronics"),
    103: ("Keyboard", "Accessories"),
    104: ("Mouse", "Accessories"),
    105: ("Printer", "Office"),
}

regions = ["Tashkent", "Samarkand", "Andijan", "Bukhara", "Fergana"]


for month in ["January", "February", "March"]:

    rows = []

    for i in range(1, 101):

        product_id = random.choice(list(products))

        product_name, category = products[product_id]

        quantity = random.randint(1, 20)
        price = random.randint(50, 1000)

        rows.append({
            "order_id": f"{month[:3]}-{i:04d}",
            "region": random.choice(regions),
            "product_id": product_id,
            "product": product_name,
            "category": category,
            "quantity": quantity,
            "price": price,
            "sales": quantity * price,
        })

    df = pd.DataFrame(rows)

    df.to_excel(
        output_dir / f"{month}.xlsx",
        index=False
    )

    print(f"{month}.xlsx created")
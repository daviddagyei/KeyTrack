import pandas as pd
import random
from faker import Faker

# Configuration
NUM_ROOMS = 15
MIN_KEYS = 3
MAX_KEYS = 10
fake = Faker()

# Column names
COLUMNS = [
    "Room Number",
    "Available Keys",
    "Collected By",
    "Lost Keys",
    "Borrowed Spare Keys",
    "Returned Keys"
]

def generate_demo_data():
    data = []
    for i in range(NUM_ROOMS):
        room_number = f"{random.randint(100, 499)}"
        total_keys = random.randint(MIN_KEYS, MAX_KEYS)

        # Decide on lost, borrowed
        lost = random.randint(0, min(2, total_keys))
        borrowed = random.randint(0, min(2, total_keys - lost))
        max_collectable = max(0, total_keys - lost - borrowed)
        if max_collectable > 0:
            collected = random.randint(1, min(3, max_collectable))
        else:
            collected = 0
        returned = random.randint(0, collected) if collected > 0 else 0

        # Available = total - (lost + borrowed + collected - returned)
        available = total_keys - (lost + borrowed + collected - returned)
        available = max(0, available)

        # Generate names for collected by
        names = [fake.name() for _ in range(collected)] if collected > 0 else []
        collected_by = ", ".join(names)

        row = [
            room_number,
            available,
            collected_by,
            lost,
            borrowed,
            returned
        ]
        data.append(row)
    return data

def main():
    data = generate_demo_data()
    df = pd.DataFrame(data, columns=COLUMNS)
    df.to_excel("key_distribution.xlsx", index=False)
    print("Demo Excel file populated with realistic data.")

if __name__ == "__main__":
    main()

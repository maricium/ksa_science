import json
from pathlib import Path

# Path to the spec file
SPEC_PATH = Path("./chemistry.json")

# Load the JSON
with open(SPEC_PATH, "r", encoding="utf-8") as f:
    spec = json.load(f)

# Flatten: topics -> subsections -> foundation_content + higher_content
concepts = []
for topic in spec.get("topics", []):
    for sub in topic.get("subsections", []):
        for c in sub.get("foundation_content", []) + sub.get("higher_content", []):
            concepts.append(c)

# Separate by tier
foundation = [c for c in concepts if c.get("tier") == "foundation"]
higher = concepts  # Higher students get everything

print("FOUNDATION CONCEPTS")
print("-" * 30)
for c in foundation:
    print(c["id"])

print("\nHIGHER CONCEPTS (all)")
print("-" * 30)
for c in higher:
    print(c["id"])

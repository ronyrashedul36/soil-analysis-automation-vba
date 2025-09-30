# --- Unique Tag Analyzer ---

print("Enter your tags, separated by commas.")
tags_input = input("Tags: ") # tech, python, coding, tech, fun


tags_list = tags_input.split(', ')

unique_tags = set(tags_list)

print("\n--- Analysis ---")
print(f"Total tags entered: {len(tags_list)}")
print(f"Number of unique tags: {len(unique_tags)}")
print(f"The unique tags are: {unique_tags}")
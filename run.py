from crawler import CoolmateCrawler

collection_urls = [
    "https://www.coolmate.me/collection/ao-ba-lo-tank-top-nam",
]

print("=== COOLMATE CRAWLER ===")
print(f"Total collections: {len(collection_urls)}")
for url in collection_urls:
    print(f"  - {url}")
print("\nStarting crawler...\n")

crawler = CoolmateCrawler(collection_urls)
crawler.run()

print("\n=== HOÀN THÀNH ===")
print("File Excel đã được lưu tại: Downloads/lecas_data.xlsx")

from crawler_tno import TheNewOriginalsCrawler

collection_urls = [
    "https://theneworiginals.co/collections/ao-thun-relaxed-fit",
]

print("=== THE NEW ORIGINALS CRAWLER ===")
print(f"Total collections: {len(collection_urls)}")
for url in collection_urls:
    print(f"  - {url}")
print("\nStarting crawler...\n")

crawler = TheNewOriginalsCrawler(collection_urls)
crawler.run()

print("\n=== HOÀN THÀNH ===")
print("File Excel đã được lưu tại: Downloads/tno_data.xlsx")

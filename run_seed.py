from seed_crawler import SeedDataCrawler

if __name__ == "__main__":
    collection_urls = [
        'https://theneworiginals.co/collections/ao-thun-relaxed-fit',
    ]
    
    print(f"=== SEED DATA CRAWLER ===")
    print(f"Will crawl {len(collection_urls)} collection(s):\n")
    for url in collection_urls:
        print(f"  - {url}")
    print()
    
    crawler = SeedDataCrawler(collection_urls)
    crawler.run()

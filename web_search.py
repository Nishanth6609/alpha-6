try:
    from ddgs import DDGS
except ModuleNotFoundError:  # pragma: no cover - optional dependency is not always installed.
    DDGS = None


def search_web(query, max_results=5):
    try:
        print("A.L.P.H.A - 6 is searching the internet...")

        if DDGS is None:
            print("Internet search unavailable: ddgs dependency is not installed.")
            return []

        with DDGS() as ddgs:
            results = list(ddgs.text(query, max_results=max_results))

        formatted_results = []
        for item in results:
            formatted_results.append({
                "title": item.get("title", "Untitled"),
                "url": item.get("href", ""),
                "description": item.get("body", "")
            })
        return formatted_results

    except Exception as error:
        print("Internet search error:", error)
        return []


if __name__ == "__main__":
    query = input("What should A.L.P.H.A - 6 search for? ")
    results = search_web(query)
    for index, result in enumerate(results, 1):
        print(f"\n{index}. {result['title']}")
        print(result["description"])
        print(result["url"])
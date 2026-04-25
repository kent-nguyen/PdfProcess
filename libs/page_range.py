def parse_pages(pages_arg):
    """Parse a pages string like '1,3,5-8,10' into a sorted list of 1-based page numbers."""
    page_set = set()
    for part in pages_arg.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            page_set.update(range(int(start), int(end) + 1))
        else:
            page_set.add(int(part))
    return sorted(page_set)

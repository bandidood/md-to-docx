#!/usr/bin/env python3

import markdown

# Test markdown content
test_md = """
# Test

Inline code: `print("hello")`

Code block:
```python
def hello():
    print("Hello, World!")
    return True
```

Another paragraph.
"""

# Convert to HTML
md = markdown.Markdown(extensions=['fenced_code', 'codehilite'])
html = md.convert(test_md)

print("Generated HTML:")
print("=" * 50)
print(html)
print("=" * 50)

# Also test with tables extension
md2 = markdown.Markdown(extensions=['tables', 'fenced_code', 'codehilite'])
html2 = md2.convert(test_md)

print("\nWith tables extension:")
print("=" * 50)
print(html2)
print("=" * 50)

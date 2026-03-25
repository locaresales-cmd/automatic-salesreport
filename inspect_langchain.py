import sys

print("Python Path:", sys.path)

try:
    import langchain
    print("langchain version:", langchain.__version__)
except ImportError:
    print("langchain not found")

try:
    import langchain.prompts
    print("langchain.prompts found")
except ImportError as e:
    print("langchain.prompts ERROR:", e)

try:
    import langchain_core.prompts
    print("langchain_core.prompts found")
except ImportError as e:
    print("langchain_core.prompts ERROR:", e)

try:
    import langchain.output_parsers
    print("langchain.output_parsers found")
    print("dir(langchain.output_parsers):", dir(langchain.output_parsers))
except ImportError as e:
    print("langchain.output_parsers ERROR:", e)

try:
    import langchain_core.output_parsers
    print("langchain_core.output_parsers found")
    print("dir(langchain_core.output_parsers):", dir(langchain_core.output_parsers))
except ImportError as e:
    print("langchain_core.output_parsers ERROR:", e)

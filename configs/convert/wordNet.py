r"""wordNet.py -- synonyms for one word, from WordNet, for EdSharp.

    python wordNet.py word

WHY WORDNET

A thesaurus is only useful when it separates MEANINGS. "Light" as the
opposite of heavy and "light" as illumination want completely different
replacements, and a flat alphabetical list of synonyms hides that.
WordNet, Princeton's freely licensed lexical database, groups words into
senses, each with its own part of speech and short definition, so the
list EdSharp shows reads as, for example:

    bright -- adjective: emitting or reflecting light readily

OUTPUT

One line per choice, ready for a list box:

    <replacement><TAB><how it reads aloud>

Special first words instead of a list:

    NOTINSTALLED   the WordNet database is not available here
    NOWORD         the database has no entry for this word

Install the database with installPdfTools.cmd in the EdSharp folder,
which installs the thesaurus alongside the PDF reader.
"""

import sys


def describe(sPos):
    """The part of speech, in the words a person would use."""
    dNames = {"n": "noun", "v": "verb", "a": "adjective", "s": "adjective", "r": "adverb"}
    return dNames.get(sPos, sPos)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    sWord = sys.argv[1].strip().lower()
    if not sWord:
        print("NOWORD")
        return 0

    try:
        from nltk.corpus import wordnet
        wordnet.synsets("test")
    except Exception:
        # Either nltk is missing or its data has not been downloaded.
        try:
            import nltk
            nltk.download("wordnet", quiet=True)
            nltk.download("omw-1.4", quiet=True)
            from nltk.corpus import wordnet
            wordnet.synsets("test")
        except Exception:
            print("NOTINSTALLED")
            return 0

    lSenses = wordnet.synsets(sWord)
    if not lSenses:
        print("NOWORD")
        return 0

    lLines = []
    setSeen = set()
    for oSense in lSenses:
        sPos = describe(oSense.pos())
        sDefinition = oSense.definition()
        for oLemma in oSense.lemmas():
            sName = oLemma.name().replace("_", " ")
            if sName.lower() == sWord:
                continue
            sKey = (sName + "|" + sDefinition).lower()
            if sKey in setSeen:
                continue
            setSeen.add(sKey)
            lLines.append(sName + "\t" + sName + " -- " + sPos + ": " + sDefinition)

    if not lLines:
        # A word can have senses but no other lemmas: report the senses
        # themselves so the definition is still available.
        for oSense in lSenses:
            lLines.append(sWord + "\t" + sWord + " -- " + describe(oSense.pos()) + ": " + oSense.definition())

    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    print("\n".join(lLines))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        print("NOTINSTALLED")
        sys.exit(0)

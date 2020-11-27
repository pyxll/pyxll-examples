"""
A more flexible VLOOKUP using Natural Language Processing
from scikit-learn.

To use in Excel add the folder this file is in to your pythonpath
in the pyxll.cfg file and add 'nlp_vlookup' to the modules list
in the same file.
"""
from sklearn import feature_extraction, metrics
from pyxll import xl_func, xl_return
from typing import List, Any
import pandas as pd


@xl_func(name="NLP.VLOOKUP")
@xl_return("dataframe<index=False, columns=False>")
def nlp_vlookup(value: str,
                table: List[List[Any]],
                col_index: int = None,
                include_score: bool = False,
                all_matches: bool = False,
                threshold: float = 0.5) -> pd.DataFrame:
    """
    Do a 'VLOOKUP' for value in table and return the value at col_index
    of the matching row.

    :param value: Value to match.
    :param table: Table of input values with the candidates in the left most column.
    :param col_index: Index of column in the table to return.
    :param include_score: Include the score in the return value.
    :param all_matches: If True return a table of all matches.
    :param threshold: Exclude matches below this threshold (0-1).
    """
    # Get the first column in the table as the list of words to match.
    words = [str(x[0]) for x in table]

    # Vectorize all the strings by creating a Bag-of-Words matrix, which extracts
    # the vocabulary from the corpus and counts how many times the words appear in each string.
    vectorizer = feature_extraction.text.CountVectorizer()
    vectors = vectorizer.fit_transform([value]+words).toarray()

    # Then, we calculate the cosine similarity, a measure based on the angle between two
    # non-zero vectors, which equals the inner product of the same vectors normalized to both
    # have length 1.
    cosine_sim = metrics.pairwise.cosine_similarity(list(vectors))

    # Get the scores for each word and put them into a DataFrame.
    scores = cosine_sim[0][1:]
    scores_df = pd.DataFrame({"score": scores}, index=words)

    # Join the input table and the scores into a single DataFrame.
    table_df = pd.DataFrame(table, index=words)
    df = table_df.join(scores_df)

    # Filter out any with a score below the threshold.
    df = df[df["score"] >= threshold]

    # If there are no matches then raise an exception.
    if not len(df.index):
        raise ValueError("No matches found")

    # Sort by score.
    df = df.sort_values(by="score", ascending=False)

    # Get the top result if not returning all matches.
    if not all_matches:
        df = df.head(1)

    # Reindex to get only the columns we're interested in and put the score first.
    columns = table_df.columns.to_list() if col_index is None else [col_index-1]
    if include_score:
        columns = ["score"] + columns
    df = df.reindex(columns=columns)

    return df


if __name__ == "__main__":
    words = [
        ["Bridgewater Associates", "United States Westport, CT", "$98,918"],
        ["Renaissance Technologies", "United States East Setauket, NY", "$70,000"],
        ["Man Group", "United Kingdom London", "$62,300"],
        ["Millennium Management", "United States New York City, NY", "$43,912"],
        ["Elliott Management", "United States New York City, NY", "$42,000"],
        ["BlackRock", "United States New York City, NY", "$39,907"],
        ["Two Sigma Investments", "United States New York City, NY", "$38,842"],
        ["The Children's Investment Fund Management", "United Kingdom London", "$35,000"],
        ["Citadel LLC", "United States Chicago, IL", "$34,340"],
        ["D.E. Shaw & Co.", "United States New York City, NY", "$34,264"],
        ["AQR Capital Management", "United States Greenwich, CT", "$32,100"],
        ["Davidson Kempner Capital Management", "United States New York City, NY", "$31,850"],
        ["Farallon Capital", "United States San Francisco, CA", "$30,000"],
        ["Baupost Group", "United States Boston, MA", "$29,100"],
        ["Marshall Wace", "United Kingdom London", "$27,800"],
        ["Capula Investment Management", "United Kingdom London", "$23,000"],
        ["Canyon Capital Advisors", "United States Los Angeles, CA", "$22,800"],
        ["Wellington Management Company", "United States Boston, MA", "$21,000"],
        ["Viking Global Investors", "United States Greenwich, CT", "$19,950"],
        ["PIMCO", "United States Newport Beach, CA", "$17,453"],
    ]

    match = nlp_vlookup("Capital", words, include_score=True, all_matches=True, threshold=0.5)
    print(match)

"""
Decision tree example using scikit-learn.

This example shows how to fit a decision tree using scikit-learn
and make predictions in an Excel spreadsheet.

This code accompanies the blog post
https://www.pyxll.com/blog/machine-learning-in-excel-with-python
"""
from pyxll import xl_func, xl_app, async_call
from sklearn.tree import DecisionTreeClassifier
from sklearn.model_selection import train_test_split
import pandas as pd
import os


_zoo_classifications = {
    1: "mammal",
    2: "bird",
    3: "reptile",
    4: "fish",
    5: "amphibian",
    6: "insect",
    7: "mollusc"
}


@xl_func("float, int, int: object")
def ml_get_zoo_tree(train_size=0.75, max_depth=5, random_state=245245):
    # Load the zoo data
    dataset = pd.read_csv(os.path.join(os.path.dirname(__file__), "data", "zoo.csv"))

    # Drop the animal names since this is not a good feature to split the data on
    dataset = dataset.drop("animal_name", axis=1)

    # Split the data into a training and a testing set
    features = dataset.drop("class", axis=1)
    targets = dataset["class"]

    train_features, test_features, train_targets, test_targets = \
        train_test_split(features, targets, train_size=train_size, random_state=random_state)

    # Train the model
    tree = DecisionTreeClassifier(criterion="entropy", max_depth=max_depth)
    tree = tree.fit(train_features, train_targets)

    # Add the feature names to the tree for use in predict function
    tree._feature_names = features.columns

    return tree


@xl_func("object tree, dict<string, int> features: var")
def ml_zoo_predict(tree, features):
    # Convert the features dictionary into a DataFrame with a single row
    features = pd.DataFrame([features], columns=tree._feature_names)

    # Get the prediction from the model
    prediction = tree.predict(features)[0]

    # Update the image in Excel
    async_call(show_image_in_excel, prediction)

    return _zoo_classifications[prediction]


@xl_func("object tree, dict<string, int> features: series", auto_resize=True)
def ml_zoo_proba(tree, features):
    # Convert the features dictionary into a DataFrame with a single row
    features = pd.DataFrame([features], columns=tree._feature_names).astype(float)

    # Get the probabilities for each class
    proba = tree.predict_proba(features)[0]

    # Construct a pandas Series as the result
    index = [_zoo_classifications[c] for c in tree.classes_]
    return pd.Series(proba, index=index)


def show_image_in_excel(classification, figname="prediction_image"):
    """Plot a figure in Excel"""
    # Show the figure in Excel as a Picture object on the same sheet
    # the function is being called from.
    xl = xl_app()
    sheet = xl.ActiveSheet

    # if a picture with the same figname already exists then get the position
    # and size from the old picture and delete it.
    for old_picture in sheet.Pictures():
        if old_picture.Name == figname:
            height = old_picture.Height
            width = old_picture.Width
            top = old_picture.Top
            left = old_picture.Left
            old_picture.Delete()
            break
    else:
        # otherwise create a new image
        top_left = sheet.Cells(1, 1)
        top = top_left.Top
        left = top_left.Left
        width, height = 100, 100

    # insert the picture
    # Ref: http://msdn.microsoft.com/en-us/library/office/ff198302%28v=office.15%29.aspx
    filename = os.path.join(os.path.dirname(__file__), "images", _zoo_classifications[classification] + ".jpg")
    picture = sheet.Shapes.AddPicture(Filename=filename,
                                      LinkToFile=0,  # msoFalse
                                      SaveWithDocument=-1,  # msoTrue
                                      Left=left,
                                      Top=top,
                                      Width=width,
                                      Height=height)

    # set the name of the new picture so we can find it next time
    picture.Name = figname

# %%
import pandas as pd

titanic = pd.read_csv(".\\titanic.csv")

titanic.head()
# %%
above_35 = titanic[titanic["Age"] > 35]
above_35.head(10)
# %%
adult_name = titanic.loc[titanic["Age"] > 35, "Name"]
print(adult_name.shape)
adult_name.head()

# %%
import pandas as pd

mydict = [{'a': 1, 'b': 2, 'c': 3, 'd': 4},
          {'a': 100, 'b': 200, 'c': 300, 'd': 400},
          {'a': 1000, 'b': 2000, 'c': 3000, 'd': 4000}]

df = pd.DataFrame(mydict)

df.iloc[[0, 2], [1, 3]]
# %%

# %%

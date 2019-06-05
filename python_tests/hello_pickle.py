import pickle

# To save
dic = {'a': 1, 'b': 2, 'c': 3}
pickle.dump(dic, open('dic.p', 'wb'))

# To open
dic = pickle.load(open('dic.p', 'rb'))

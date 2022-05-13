def find_in_list(string_list, key_word):
    """
    Gives an index number for the given element in the given list.

    :param string_list: A list
    :param key_word: element to be found
    :return: the index number of the element
    :raises ValueError: if the element is not found
    """
    try:
        return string_list.index(key_word)
    except ValueError:
        return None

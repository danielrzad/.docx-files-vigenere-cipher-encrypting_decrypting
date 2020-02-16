from docx import Document


def vigenère_cipher_encrypting(message, characters_matrix, keyword):
    letters = (
    	'1234567890-=!@#$%^&*()_+'
    	'AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz'
	)
    encrypting_key = ''
    c = 0
    while len(encrypting_key) != len(message):
        encrypting_key += keyword[c]
        if c + 1 > len(keyword) - 1:
            c = 0
        else:
            c += 1
    encrypted_message = ''
    for idx, msg_l in enumerate(message):
        #msg_l = message letter at given index(idx)
        if msg_l in letters:
            encrypted_message += letters[
            (letters.find(msg_l) 
            + letters.find(encrypting_key[idx]))
            % len(letters)
			]
        else:
            encrypted_message += msg_l
    return encrypted_message


def vigenère_cipher_decrypting(message, characters_matrix, keyword):
    letters = (
    	'1234567890-=!@#$%^&*()_+'
    	'AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz'
	)
    encrypting_key = ''
    c = 0
    while len(encrypting_key) != len(message):
        encrypting_key += keyword[c]
        if c + 1 > len(keyword) - 1:
            c = 0
        else:
            c += 1
    decrypted_message = ''
    for idx, msg_l in enumerate(message):
        if msg_l in letters:
            decrypted_message += letters[
            (letters.find(msg_l)
            - letters.find(encrypting_key[idx]))
            % len(letters)
			]
        else:
            decrypted_message += msg_l
    return decrypted_message


def main(characters_matrix, keyword, docx_location, docx_name, encrypt, decrypt):
    
    """
                            !IMPORTANT!
    Once encrypted file will not be decryptable without 
    characters_matrix and keyword given during encrypting process.
    The same characters_matrix and keyword needs to be used in process
    of encrypting or decrypting!

    characters_matrix = long string of letters for eg. Alphabet list
    keyword = your secret key which will be used to code/decrypt file
    docx_location = location of MS .docx file which u want to 
        encrypt/decrypt
    code = set to True if You want to code MS .docx 
        if You want to decrypt the file please set to False
    decrypt = set to True if You want to decrypt MS .docx 
        if You want to code the file please set to False
    """

    document = Document(f'{docx_location}{docx_name}')

    for paragraph in document.paragraphs:
        if code == True:
            paragraph.text = vigenère_cipher_encrypting(
                paragraph.text, 
                characters_matrix, 
                keyword,
			)
        else:
            paragraph.text = vigenère_cipher_decrypting(
                paragraph.text, 
                characters_matrix, 
                keyword,
			)

    document.save(f'{docx_location}{docx_name}')


if __name__ == '__main__':
    main(
        characters_matrix='old/', 
        keyword='friends',
        docx_location='C:/Users/Admin/Desktop/Python/to_code/',
        docx_name='lorem_ipsum.docx',
        code=False,
        decrypt=True,
	)
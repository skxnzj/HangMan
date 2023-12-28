import pandas as pd
import random

def load_words(file_path):
    df = pd.read_excel(file_path)
    return list(df['Parole'])

def choose_random_word(words_list):
    return random.choice(words_list)

def display_word(word, guessed_letters):
    return ''.join(letter if letter in guessed_letters else '_' for letter in word)

def main():
    excel_file = 'Parole.xlsx'
    words_list = load_words(excel_file)
    selected_word = choose_random_word(words_list)

    print(selected_word)

    max_attempts = len(selected_word) + 2
    attempts = 0
    guessed_letters = set()

    print("Welcome to the Word Guessing Game!")
    print("Try to guess the word.")

    while attempts < max_attempts:
        current_display = display_word(selected_word, guessed_letters)
        print(f"\nCurrent word: {current_display}")

        if '_' not in current_display:
            print("Congratulations! You guessed the word.")
            break

        print(f"Attempts remaining: {max_attempts - attempts}")
        guess = input("Enter a letter: ").lower()

        if guess.isalpha() and len(guess) == 1:
            if guess in guessed_letters:
                print("You've already guessed that letter. Try again.")
            elif guess in selected_word:
                print("Good guess!")
                guessed_letters.add(guess)
            else:
                print("Incorrect guess. Try again.")
                attempts += 1
        else:
            print("Invalid input. Please enter a single letter.")

    if attempts == max_attempts:
        print(f"\nSorry, you've run out of attempts. The word was: {selected_word}")

if __name__ == "__main__":
    main()

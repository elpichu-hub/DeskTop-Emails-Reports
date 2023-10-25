import os
import openai
openai.api_key = os.getenv("OPENAI_API_KEY")


def call_gpt_api(content):
    completion = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "system", "content": "You are my assistant, an expert in all sort of matters"},
        {"role": "user", "content": content}
    ]
    )

    # print(completion.choices[0].message)
    return completion.choices[0].message

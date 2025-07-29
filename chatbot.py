def get_response(message):
    # Simple rule-based response for demo
    if 'hello' in message.lower():
        return 'Hello! How can I help you today?'
    elif 'doc' in message.lower():
        return 'I can generate documentation from our chat.'
    else:
        return 'Tell me more about what you need.'

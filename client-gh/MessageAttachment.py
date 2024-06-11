class MessageAttachment:
    def __init__(self, internet_message_id, source_message_id, has_attachments):
        self.internet_message_id = internet_message_id
        self.source_message_id = source_message_id
        self.has_attachments = has_attachments
        
    def get_message_id_(internet_message_id, message_creation_responses):
        for created_msg in message_creation_responses:
            if created_msg.internet_message_id == internet_message_id:
                return created_msg.id
        return None
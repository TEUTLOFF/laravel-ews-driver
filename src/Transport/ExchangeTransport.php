<?php

namespace Adeboyed\LaravelExchangeDriver\Transport;

use Illuminate\Mail\Transport\Transport;
use Illuminate\Support\Arr;
use jamesiarmes\PhpEws\ArrayType\ArrayOfRecipientsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAttachmentsType;
use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;
use jamesiarmes\PhpEws\Request\CreateItemType;
use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\FileAttachmentType;
use jamesiarmes\PhpEws\Type\MessageType;
use jamesiarmes\PhpEws\Type\SingleRecipientType;
use Swift_Attachment;
use Swift_Mime_SimpleMessage;

class ExchangeTransport extends Transport
{
    protected $host;

    protected $username;

    protected $password;

    protected $messageDispositionType;

    public function __construct($host, $username, $password, $messageDispositionType)
    {
        $this->host = $host;
        $this->username = $username;
        $this->password = $password;
        $this->messageDispositionType = $messageDispositionType;
    }

    public function send(Swift_Mime_SimpleMessage $message, &$failedRecipients = null)
    {
        $this->beforeSendPerformed($message);

        $client = new Client(
            $this->host,
            $this->username,
            $this->password
        );

        $request = new CreateItemType();
        $request->Items = new NonEmptyArrayOfAllItemsType();

        $request->MessageDisposition = $this->messageDispositionType;

        // Create the ewsMessage.
        $ewsMessage = new MessageType();
        $ewsMessage->Subject = $message->getSubject();

        // Set the sender.
        $ewsMessage->From = new SingleRecipientType();
        $ewsMessage->From->Mailbox = new EmailAddressType();
        $ewsMessage->From->Mailbox->EmailAddress = config('mail.from.address');

        $toRecipients = $message->getTo();
        $ccRecipients = $message->getCc();
        $bccRecipients = $message->getBcc();
        $replyToRecipients = $message->getReplyTo();

        $ewsMessage->ToRecipients = $this->createArrayOfRecipientsType($toRecipients);
        $ewsMessage->CcRecipients = $this->createArrayOfRecipientsType($ccRecipients);
        $ewsMessage->BccRecipients = $this->createArrayOfRecipientsType($bccRecipients);
        $ewsMessage->ReplyTo = $this->createArrayOfRecipientsType($replyToRecipients);

        // Set the ewsMessage body.
        $ewsMessage->Body = new BodyType();
        $ewsMessage->Body->BodyType = BodyTypeType::HTML;
        $ewsMessage->Body->_ = $message->getBody();

        // Attachments
        if ($attachments = $message->getChildren()) {
            $ewsMessage->Attachments = new NonEmptyArrayOfAttachmentsType();

            foreach ($attachments as $attachment) {
                if ($attachment instanceof Swift_Attachment) {
                    $fileAttachment = new FileAttachmentType();

                    $fileAttachment->Name = $attachment->getFilename();
                    $fileAttachment->Content = base64_encode($attachment->getBody());
                    $fileAttachment->ContentType = $attachment->getContentType();

                    $ewsMessage->Attachments->FileAttachment[] = $fileAttachment;
                }
            }
        }

        $request->Items->Message[] = $ewsMessage;
        $response = $client->CreateItem($request);

        // Iterate over the results, printing any error messages or ewsMessage ids.
        $response_messages = $response->ResponseMessages->CreateItemResponseMessage;
        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $ewsMessage
                    = $response_message->MessageText;
                fwrite(STDERR, "Message failed to create with \"{$code}: {$ewsMessage}\"\n");
            }
        }

        $this->sendPerformed($message);

        return $this->numberOfRecipients($message);
    }

    private function createArrayOfRecipientsType(string|array $recipients): ArrayOfRecipientsType
    {
        $arrayOfRecipientsType = new ArrayOfRecipientsType();

        foreach (Arr::wrap($recipients) as $email => $name) {
            $arrayOfRecipientsType->Mailbox[] = $this->getRecipient($email, $name);
        }

        return $arrayOfRecipientsType;
    }

    private function getRecipient(int|string $email, mixed $name): EmailAddressType
    {
        $recipient = new EmailAddressType();
        $recipient->EmailAddress = $email;

        if ($name != null) {
            $recipient->Name = $name;
        }

        return $recipient;
    }
}

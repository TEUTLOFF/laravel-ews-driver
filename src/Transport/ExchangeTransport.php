<?php

namespace Adeboyed\LaravelExchangeDriver\Transport;

use Illuminate\Support\Arr;
use Illuminate\Support\Facades\Log;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAttachmentsType;
use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\MessageDispositionType;
use jamesiarmes\PhpEws\Request\CreateItemType;

use jamesiarmes\PhpEws\ArrayType\ArrayOfRecipientsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;

use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;

use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\FileAttachmentType;
use jamesiarmes\PhpEws\Type\MessageType;
use jamesiarmes\PhpEws\Type\SingleRecipientType;
use Symfony\Component\Mailer\Envelope;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\TransportInterface;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\RawMessage;

class ExchangeTransport implements TransportInterface
{
    protected $host;
    protected $username;
    protected $password;
    protected $messageDispositionType;
    protected $client;

    public function __construct($host, $username, $password, $messageDispositionType)
    {
        $this->host = $host;
        $this->username = $username;
        $this->password = $password;
        $this->messageDispositionType = !empty($messageDispositionType)
            ? $messageDispositionType
            : MessageDispositionType::SEND_ONLY;
    }

    public function __toString(): string
    {
        return "exchange";
    }

    public function setClient(Client $client): void
    {
        $this->client = $client;
    }

    public function send(RawMessage $message, ?Envelope $envelope = null): ?SentMessage
    {
        if (!($message instanceof Email)) {
            throw new \InvalidArgumentException(
                sprintf('"%s" only supports instances of "%s" (instance of "%s" given).',
                    __CLASS__,
                    Email::class,
                    get_debug_type($message)
                )
            );
        }

        return $this->sendInternal($message);
    }

    private function sendInternal(Email $message, ?Envelope $envelope = null): ?SentMessage
    {
        if (!$this->client) {
            $this->client = new Client(
                $this->host,
                $this->username,
                $this->password
            );
        }

        $request = new CreateItemType();
        $request->Items = new NonEmptyArrayOfAllItemsType();

        $request->MessageDisposition = $this->messageDispositionType;

        // Create the ewsMessage.
        $ewsMessage = new MessageType();
        $ewsMessage->Subject = $message->getSubject();

        // Set the sender.
        $ewsMessage->From = new SingleRecipientType();
        $ewsMessage->From->Mailbox = new EmailAddressType();
        $ewsMessage->From->Mailbox->EmailAddress = $message->getFrom()[0]->getAddress();
        $ewsMessage->From->Mailbox->Name = $message->getFrom()[0]->getName();

        // Set the recipient.
        $ewsMessage->ToRecipients = new ArrayOfRecipientsType();
        $ewsMessage->ToRecipients->Mailbox = [];

        // Add all recipients.
        $ewsMessage->ToRecipients = $this->createArrayOfRecipientsType($message->getTo());
        $ewsMessage->CcRecipients = $this->createArrayOfRecipientsType($message->getCc());
        $ewsMessage->BccRecipients = $this->createArrayOfRecipientsType($message->getBcc());
        $ewsMessage->ReplyTo = $this->createArrayOfRecipientsType($message->getReplyTo());

        // Set the ewsMessage body.
        $ewsMessage->Body = new BodyType();
        $ewsMessage->Body->BodyType = BodyTypeType::HTML;
        $ewsMessage->Body->_ = $message->getHtmlBody();

        // Attachments
        if ($attachments = $message->getAttachments()) {
            $ewsMessage->Attachments = new NonEmptyArrayOfAttachmentsType();

            foreach ($attachments as $attachment) {
                $fileAttachment = new FileAttachmentType();

                $fileAttachment->Name = $attachment->getFilename();
                $fileAttachment->Content = $attachment->getBody();
                $fileAttachment->ContentType = $attachment->getContentType();

                $ewsMessage->Attachments->FileAttachment[] = $fileAttachment;
            }
        }

        $request->Items->Message[] = $ewsMessage;
        $response = $this->client->CreateItem($request);

        // Iterate over the results, printing any error messages or ewsMessage ids.
        $response_messages = $response->ResponseMessages->CreateItemResponseMessage;
        foreach ($response_messages as $response_message) {
            // Make sure the request succeeded.
            if ($response_message->ResponseClass != ResponseClassType::SUCCESS) {
                $code = $response_message->ResponseCode;
                $ewsMessage
                    = $response_message->MessageText;
                Log::error("Message failed to create with \"$code: $ewsMessage\"\n");
            }
        }

        if ($envelope === null) {
            $envelope = new Envelope(
                $message->getFrom()[0],
                $message->getTo()
            );
        }

        return new SentMessage($message, $envelope);
    }

    private function createArrayOfRecipientsType(Address|array|null $recipients): ArrayOfRecipientsType
    {
        $arrayOfRecipientsType = new ArrayOfRecipientsType();

        foreach (Arr::wrap($recipients) as $recipient) {
            $arrayOfRecipientsType->Mailbox[] = $this->getRecipient($recipient);
        }

        return $arrayOfRecipientsType;
    }

    private function getRecipient(Address $to): EmailAddressType
    {
        $recipient = new EmailAddressType();

        $recipient->EmailAddress = $to->getAddress();
        $recipient->Name = $to->getName();

        return $recipient;
    }
}

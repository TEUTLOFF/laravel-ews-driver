<?php

namespace Adeboyed\LaravelExchangeDriver\Transport;

use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Request\CreateItemType;

use jamesiarmes\PhpEws\ArrayType\ArrayOfRecipientsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;

use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;

use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\MessageType;
use jamesiarmes\PhpEws\Type\SingleRecipientType;
use Symfony\Component\Mailer\Envelope;
use Symfony\Component\Mailer\MailerInterface;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\TransportInterface;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\RawMessage;

class ExchangeTransport implements TransportInterface
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

    public function __toString(): string
    {
        return "exchange";
    }

    public function send(RawMessage $message, ?Envelope $envelope = null): ?SentMessage
    {
        if (!($message instanceof Email)) {
            throw new \InvalidArgumentException(sprintf('"%s" transport only supports instances of "%s" (instance of "%s" given).', __CLASS__, Email::class, get_debug_type($message)));
        }

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
        $ewsMessage->ToRecipients = new ArrayOfRecipientsType();

        // Set the sender.
        $ewsMessage->From = new SingleRecipientType();
        $ewsMessage->From->Mailbox = new EmailAddressType();
        $ewsMessage->From->Mailbox->EmailAddress = config('mail.from.address');

        // Set the recipient.
        $ewsMessage->ToRecipients = new ArrayOfRecipientsType();
        $ewsMessage->ToRecipients->Mailbox = [];

        // Add all recipients.
        foreach ($message->getTo() as $rec) {
            $recipient = new EmailAddressType();
            $recipient->EmailAddress = $rec->getAddress();
            $ewsMessage->ToRecipients->Mailbox[] = $recipient;
        }

        // Set the ewsMessage body.
        $ewsMessage->Body = new BodyType();
        $ewsMessage->Body->BodyType = BodyTypeType::HTML;
        $ewsMessage->Body->_ = $message->getHtmlBody();

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
                fwrite(STDERR, "Message failed to create with \"$code: $ewsMessage\"\n");
                continue;
            }
        }

        return new SentMessage($message, $envelope);
    }
}

<?php

namespace Adeboyed\LaravelExchangeDriver\Transport;

use Illuminate\Mail\Transport\Transport;
use jamesiarmes\PhpEws\ArrayType\ArrayOfRecipientsType;
use jamesiarmes\PhpEws\ArrayType\NonEmptyArrayOfAllItemsType;
use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\BodyTypeType;
use jamesiarmes\PhpEws\Enumeration\ResponseClassType;
use jamesiarmes\PhpEws\Request\CreateItemType;
use jamesiarmes\PhpEws\Type\BodyType;
use jamesiarmes\PhpEws\Type\EmailAddressType;
use jamesiarmes\PhpEws\Type\MessageType;
use jamesiarmes\PhpEws\Type\SingleRecipientType;
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
        $ewsMessage->ToRecipients = new ArrayOfRecipientsType();

        // Set the sender.
        $ewsMessage->From = new SingleRecipientType();
        $ewsMessage->From->Mailbox = new EmailAddressType();
        $ewsMessage->From->Mailbox->EmailAddress = config('mail.from.address');

        // Set the recipient.
        foreach ($this->allContacts($message) as $email => $name) {
            $recipient = new EmailAddressType();
            $recipient->EmailAddress = $email;
            if ($name != null) {
                $recipient->Name = $name;
            }
            $ewsMessage->ToRecipients->Mailbox[] = $recipient;
        }

        // Set the ewsMessage body.
        $ewsMessage->Body = new BodyType();
        $ewsMessage->Body->BodyType = BodyTypeType::HTML;
        $ewsMessage->Body->_ = $message->getBody();

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

                continue;
            }
        }

        $this->sendPerformed($message);

        return $this->numberOfRecipients($message);
    }

    /**
     * Get all of the contacts for the ewsMessage
     *
     * @param  \Swift_Mime_SimpleMessage  $ewsMessage
     * @return array
     */
    protected function allContacts(Swift_Mime_SimpleMessage $message)
    {
        return array_merge(
            (array) $message->getTo(),
            (array) $message->getCc(),
            (array) $message->getBcc()
        );
    }
}

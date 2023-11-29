<?php

namespace Adeboyed\LaravelExchangeDriver\Transport;

use jamesiarmes\PhpEws\Client;
use jamesiarmes\PhpEws\Enumeration\MessageDispositionType;
use PHPUnit\Framework\TestCase;
use stdClass;
use Symfony\Component\Mailer\Exception\TransportExceptionInterface;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\Part\DataPart;

class ExchangeTransportTest extends TestCase
{
    /**
     * @throws TransportExceptionInterface
     */
    public function testSendInternalWithMockedClient()
    {
        // Mock the Client
        $clientMock = $this->createMock(Client::class);

        $responseMessage = new stdClass();
        $responseMessage->ResponseClass = 'Success'; // or 'Error' based on what you need to test
        $responseMessage->ResponseCode = 'NoError';  // or an error code
        $responseMessage->MessageText = 'Message text'; // or an error message

        $responseMessages = [$responseMessage];

        $response = new stdClass();
        $response->ResponseMessages = new stdClass();
        $response->ResponseMessages->CreateItemResponseMessage = $responseMessages;
        $clientMock->method('CreateItem')->willReturn($response);
        $transport = new ExchangeTransport('host', 'username', 'password', 'DispositionType');
        $transport->setClient($clientMock);

        $email = (new Email())
            ->from('sender@example.com')
            ->to('recipient@example.com')
            ->cc('cc@example.com')
            ->bcc('bcc@example.com')
            ->replyTo('replyto@example.com')
            ->subject('Test Subject')
            ->html('<p>Test Body</p>');

        $attachmentContent = 'asdf';
        $email->attach($attachmentContent, 'example.txt', 'text/plain');

        $result = $transport->send($email);
        $this->assertInstanceOf(SentMessage::class, $result);

        $emailAsString = $result->toString();
        $this->assertStringContainsString('From: sender@example.com', $emailAsString);
        $this->assertStringContainsString('To: recipient@example.com', $emailAsString);
        $this->assertStringContainsString('Cc: cc@example.com', $emailAsString);
        $this->assertStringContainsString('Reply-To: replyto@example.com', $emailAsString);
        $this->assertStringContainsString('Subject: Test Subject', $emailAsString);
        $this->assertStringContainsString('<p>Test Body</p>', $emailAsString);
        $this->assertStringContainsString('Content-Type: text/plain; name=example.txt', $emailAsString);
        $this->assertStringContainsString('Content-Disposition: attachment; name=example.txt; filename=example.txt', $emailAsString);
        $this->assertStringContainsString('YXNkZg==', $emailAsString);
    }
}

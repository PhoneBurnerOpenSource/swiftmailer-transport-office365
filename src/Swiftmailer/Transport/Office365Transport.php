<?php

namespace PhoneburnerOpenSource\Office365Swiftmailer\Swiftmailer\Transport;

use GuzzleHttp\Exception\ClientException;
use Microsoft\Graph\Graph;
use Microsoft\Graph\Model\BodyType;
use Microsoft\Graph\Model\EmailAddress;
use Microsoft\Graph\Model\FileAttachment;
use Microsoft\Graph\Model\ItemBody;
use Microsoft\Graph\Model\Message;
use Microsoft\Graph\Model\Recipient;
use PhoneburnerOpenSource\Office365Swiftmailer\Swiftmailer\Exception\InvalidToken;
use PhoneburnerOpenSource\Office365Swiftmailer\Swiftmailer\Exception\Office365TransportException;
use Swift_DependencyContainer;
use Swift_Events_EventDispatcher;
use Swift_Events_EventListener;
use Swift_Events_SendEvent;
use Swift_Mime_SimpleMessage;
use Swift_Transport;

class Office365Transport implements Swift_Transport
{
    /**
     * @var bool
     */
    protected $started = false;

    protected $eventDispatcher;

    /**
     * @var Graph
     */
    protected $client;

    public function __construct(string $access_token, ?Swift_Events_EventDispatcher $eventDispatcher = null)
    {
        if ($eventDispatcher === null) {
            $container = Swift_DependencyContainer::getInstance();
            $eventDispatcher = $container->lookup('transport.eventdispatcher');
        }
        $this->eventDispatcher = $eventDispatcher;

        $this->client = new Graph();
        $this->client->setAccessToken($access_token);
    }

    public function isStarted(): bool
    {
        return $this->started;
    }

    public function start(): void
    {
        $this->started = true;
    }

    public function stop(): void
    {
        $this->started = false;
    }

    public function ping(): bool
    {
        return true;
    }

    public function send(Swift_Mime_SimpleMessage $message, &$failedRecipients = null): int
    {
        if ($evt = $this->eventDispatcher->createSendEvent($this, $message)) {
            $this->eventDispatcher->dispatchEvent($evt, 'beforeSendPerformed');
            if ($evt->bubbleCancelled()) {
                return 0;
            }
        }

        $from = $message->getFrom();
        reset($from);

        $email = new Message();
        $sender_address = new EmailAddress();
        $sender_address->setAddress(key($from));
        $sender_address->setName(current($from));

        $count = 0;

        if ($to = $message->getTo()) {
            $recipients = [];
            foreach ($to as $address => $name) {
                $email_address = new EmailAddress();
                $email_address->setAddress($address);
                $email_address->setName($name);
                $recipient = new Recipient();
                $recipient->setEmailAddress($email_address);
                $recipients[] = $recipient;
                $count++;
            }
            $email->setToRecipients($recipients);
        }

        if ($bcc = $message->getBcc()) {
            $recipients = [];
            foreach ($bcc as $address => $name) {
                $count++;

                $email_address = new EmailAddress();
                $email_address->setAddress($address);
                $email_address->setName($name);
                $recipient = new Recipient();
                $recipient->setEmailAddress($email_address);
                $recipients[] = $recipient;
            }
            $email->setBccRecipients($recipients);
        }

        if ($cc = $message->getCc()) {
            foreach ($cc as $address => $name) {
                $count++;

                $email_address = new EmailAddress();
                $email_address->setAddress($address);
                $email_address->setName($name);
                $recipient = new Recipient();
                $recipient->setEmailAddress($email_address);
                $recipients[] = $recipient;
            }
            $email->setCcRecipients($recipients);
        }

        $email->setSubject($message->getSubject());

        $body = new ItemBody();

        $body->setContent($message->getBody());
        $body_type = new BodyType(BodyType::TEXT);
        if (strtolower($message->getBodyContentType()) === 'text/html') {
            $body_type = new BodyType(BodyType::HTML);
        }
        $body->setContentType($body_type);

        $children = $message->getChildren();
        $attachments = [];
        foreach ($children as $child) {
            $headers = $child->getHeaders();
            $content = $headers->get('Content-Disposition');
            if ($content) {
                $attachment = new FileAttachment();
                $attachment->setName($content->getParameter('filename'));
                $attachment->setContentType($child->getBodyContentType());
                $attachment->setContentBytes($child->getBody());
                $attachment->setODataType("#microsoft.graph.fileAttachment");

                $attachments[] = $attachment;
            }
        }

        if ($attachments) {
            $email->setAttachments($attachments);
        }

        $email->setBody($body);
        $body = ['message' => $email];
        try {
            $response = $this->client->createRequest('POST', '/me/sendmail')
                ->attachBody($body)
                ->execute();
        } catch (ClientException $e) {
            if ($e->getCode() === 401) {
                $data = json_decode($e->getResponse()->getBody(), true);
                throw new InvalidToken($data['error']['code'] ?? 'Invalid Token');
            }
            throw $e;
        }


        if (202 === $response->getStatus()) {
            if ($evt) {
                $evt->setResult(Swift_Events_SendEvent::RESULT_SUCCESS);
                $evt->setFailedRecipients($failedRecipients);
                $this->eventDispatcher->dispatchEvent($evt, 'sendPerformed');
            }

            return $count;
        }

        $this->throwException(
            new Office365TransportException('Response error: ' . $response->statusCode(), 0, null, $response)
        );
    }

    public function registerPlugin(Swift_Events_EventListener $plugin): void
    {
        $this->eventDispatcher->bindEventListener($plugin);
    }

    private function throwException(Office365TransportException $e)
    {
        if ($evt = $this->eventDispatcher->createTransportExceptionEvent($this, $e)) {
            $this->eventDispatcher->dispatchEvent($evt, 'exceptionThrown');
            if ( ! $evt->bubbleCancelled()) {
                throw $e;
            }
        } else {
            throw $e;
        }
    }
}

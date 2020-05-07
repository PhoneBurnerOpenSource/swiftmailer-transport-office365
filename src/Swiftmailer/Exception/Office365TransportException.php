<?php

namespace PhoneburnerOpenSource\Office365Swiftmailer\Swiftmailer\Exception;

use Exception;
use Microsoft\Graph\Http\GraphResponse;
use Swift_TransportException;

class Office365TransportException extends Swift_TransportException
{
    /**
     * @var GraphResponse|null
     */
    protected $response;

    public function __construct(string $message, int $code = 0, ?Exception $previous = null, ?GraphResponse $response = null)
    {
        parent::__construct($message, $code, $previous);
        $this->response = $response;
    }

    public function getResponse(): ?GraphResponse
    {
        return $this->response;
    }
}

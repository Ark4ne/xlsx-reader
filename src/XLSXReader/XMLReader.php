<?php


namespace Ark4ne\XLSXReader;

/**
 * Class XMLReader
 * @package Ark4ne\XLSXReader
 */
class XMLReader extends \XMLReader
{
    /**
     * Move cursor to node by his name.
     *
     * @param string $node_name
     *
     * @return bool
     */
    public function find(string $node_name): bool
    {
        /**
         * @noinspection PhpStatementHasEmptyBodyInspection
         * @noinspection PhpStatementWithoutBracesInspection
         */
        while (
            $this->read()
            && $this->nodeType !== self::NONE
            && ($this->nodeType !== self::ELEMENT || $this->name !== $node_name)
        ) /** @noinspection SuspiciousSemicolonInspection */ ;

        return $this->name === $node_name;
    }

    public function readValue()
    {
        $current_type = $this->nodeType;

        if ($current_type === self::TEXT) {
            return $this->value;
        }

        if ($current_type !== self::ELEMENT) {
            return null;
        }

        $depth = 0;

        $content = '';

        while ($this->read()) {
            $current_type = $this->nodeType;

            if ($current_type === self::TEXT || $current_type === self::WHITESPACE) {
                $content .= $this->value;
            } elseif ($current_type === self::ELEMENT) {
                $depth++;
            } elseif ($current_type === self::END_ELEMENT) {
                if ($depth === 0) {
                    return $content;
                }

                $depth--;
            }
        }

        return null;
    }
}

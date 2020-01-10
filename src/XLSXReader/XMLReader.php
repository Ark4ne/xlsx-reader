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
        while (
            $this->read()
            && $this->nodeType !== self::NONE
            && ($this->nodeType !== self::ELEMENT || $this->name !== $node_name)
        ) ;

        return $this->name === $node_name;
    }
}

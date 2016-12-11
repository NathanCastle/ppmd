#!/bin/bash

#script to recursively traverse a dir

function traverse() {
for file in "$1"/*
do
    if [ ! -d "${file}" ] ; then
        if [[ $file =~ \.pptx$ ]]; then
          python3 main.py $file
        fi
    else
        traverse "${file}"
    fi
done
}

function main() {
    traverse "$1"
}

main "$1"
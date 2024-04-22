#!/bin/bash

list_drives() {
    echo "Available drives:"
    lsblk -d -o NAME,SIZE | grep -Ev 'loop|sr0|$(mount | grep " / " | cut -d" " -f1)'
}

select_drive() {
    local prompt="$1"
    local drive

    while [[ -z "$drive" ]]; do
        read -p "$prompt: " drive
        if [[ ! -b "/dev/$drive" ]]; then
            echo "Invalid drive. Please enter a valid drive name."
            drive=""
        fi
    done

    echo "/dev/$drive"
}

get_drive_size() {
    local drive="$1"
    lsblk -bdno SIZE "/dev/$drive"
}

check_drive_size() {
    local source_drive="$1"
    local destination_drive="$2"

    local source_size=$(get_drive_size "$source_drive")
    local destination_size=$(get_drive_size "$destination_drive")

    if (( source_size >= destination_size )); then
        echo "Error: Source drive is larger than the destination drive."
        return 1
    fi

    return 0
}

echo "Listing available drives:"
list_drives

echo "Select the source drive:"
source_drive=$(select_drive "Enter the name of the source drive")

echo "Select the destination drive:"
destination_drive=$(select_drive "Enter the name of the destination drive")

echo "Source drive:        $source_drive"
echo "Destination drive:   $destination_drive"

read -p "Are you sure you want to proceed? (y/n): " confirm

if [[ "$confirm" == "y" || "$confirm" == "Y" ]]; then
    echo "Cloning source to destination and calculating MD5 checksum of source..."
    sudo dd if="$source_drive" bs=4M status=progress conv=sync,noerror | tee >(md5sum > source_drive.md5)
    md5source=$(cat source_drive.md5 | cut -d" " -f1)
    echo "Calculating destination MD5 checksum..."
    md5destination=$(sudo dd if="$destination_drive" bs=4M status=progress count=$(( $(get_drive_size "$source_drive") / 4194304 )) | md5sum | cut -d" " -f1)
    echo "MD5 checksum of source:      $md5source"
    echo "MD5 checksum of destination: $md5destination"
    if [[ "$md5source" == "$md5destination" ]]; then
        echo "MD5 checksums match. Cloning completed successfully."
        echo "Zeroing out unused space on destination drive..."
        if sudo dd if=/dev/zero of="$destination_drive" bs=4M status=progress conv=fdatasync,notrunc seek=$(( $(get_drive_size "$source_drive") / 4194304 )) && sync; then
            echo "Zeroing out unused space completed."
        else
            echo "Error: Failed to zero out unused space on destination drive."
            exit 1
        fi
    else
        echo "ERROR: MD5 checksums do not match. Cloning may have failed."
    fi
    echo "Cloning and zeroing out unused space completed."
else
    echo "Operation aborted."
fi

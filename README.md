# ESXiV
This installs the ESXi 6.0.0 (with net-tulip driver) on Hyper-V

## How to use

* Use the online bundle
```
ESXiV -Online -VMHome <Path where all the files of the VM stores>
```
* Use the downloaded offline bundle  
*note that the bundle should be or older than **ESXi-6.0.0-20170604001***
```
ESXiV -Bundle <Path to the Bundle> -VMHome <Path where all the files of the VM stores>
```

## Other options
* `-Pswd <password>` *specifies the `password` of the EXSi server's root*
* `-VMName <name>` *specifies the name of the vm on Hyper-V*

## Known limitations
* AFAIK, The only working network switch is the `external switch`, which will not allow host to vm connections when you only have one nic. See https://serverfault.com/questions/984241/hyper-v-connecting-to-host-to-vm-w-external-switch. 
*But you can access your esxi vm via `IPv6` instead in that case.*

## Thanks to
https://www.nakivo.com/blog/install-esxi-hyper-v  
https://communities.vmware.com/thread/511875  
https://communities.vmware.com/thread/427502  

{
  inputs = {
    google-apps-script-devshell.url = "https://api.mynixos.com/pveierland/google-apps-script-devshell/archive/main.tar.gz";
  };
  outputs = inputs@{ self, google-apps-script-devshell, ... }:
    let
      filterAttrs = pred: set: builtins.listToAttrs (builtins.concatMap
        (name: let v = set.${name}; in if pred name v then [ ((name: value: { inherit name value; }) name v) ] else [ ])
        (builtins.attrNames set));
      forwardFlakeOutputs = input: filterAttrs (n: v: !(builtins.elem n [ "inputs" "outputs" "narHash" "outPath" "sourceInfo" ])) input;
    in
    forwardFlakeOutputs google-apps-script-devshell;
}

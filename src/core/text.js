export const $=(selector,root=document)=>root.querySelector(selector);
export const $$=(selector,root=document)=>[...root.querySelectorAll(selector)];
export const esc=value=>String(value==null?'':value).replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
export const clean=value=>String(value==null?'':value).trim();
export const NAT_COLLATOR=new Intl.Collator(undefined,{numeric:true,sensitivity:'base'});
export const natCmp=(left,right)=>NAT_COLLATOR.compare(String(left),String(right));
